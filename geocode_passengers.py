"""
geocode_passengers.py

Expand passenger address fields into one record per usable address, geocode
those addresses with the U.S. Census Batch Geocoder, and create a report-ready
CSV for later evacuation-zone analysis.

The module does NOT determine evacuation zones. It leaves evacuation fields
blank so a later spatial-analysis module can populate them.

Input expected:
    spedsta_passengers.csv

Required columns:
    profile_id
    firstname
    lastname
    address 1
    address 2

Dependency:
    pip install requests

Examples:
    python geocode_passengers.py spedsta_passengers.csv
    python geocode_passengers.py spedsta_passengers.csv -o passenger_evacuation_report.csv
    python geocode_passengers.py spedsta_passengers.csv --prepare-only

Privacy note:
    This sends address text, but not passenger names, phone numbers, email
    addresses, or profile IDs, to the U.S. Census Geocoder. Temporary random
    record IDs are used in the geocoding request.
"""

from __future__ import annotations

import argparse
import csv
import io
import logging
import re
import secrets
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

import requests


LOGGER = logging.getLogger(__name__)

CENSUS_BATCH_URL = (
    "https://geocoding.geo.census.gov/geocoder/locations/addressbatch"
)
DEFAULT_BENCHMARK = "Public_AR_Current"
DEFAULT_TIMEOUT_SECONDS = 180

REQUIRED_COLUMNS = {
    "profile_id",
    "firstname",
    "lastname",
    "address 1",
    "address 2",
    "mobilephone_number",
    "homephone_number",
}

OUTPUT_COLUMNS = [
    "passenger_name",
    "profile_id",
    "mobile_phone",
    "home_phone",
    "phone_numbers",
    "address_source",
    "input_address",
    "street",
    "city",
    "state",
    "zip",
    "geocode_status",
    "match_type",
    "matched_address",
    "longitude",
    "latitude",
    "tigerline_id",
    "tigerline_side",
    "evacuation_zone",
    "proximity_to_zone_miles",
    "proximity_status",
    "review_notes",
]

INVALID_ADDRESS_PHRASES = {
    "incorrect mailing address on file",
    "no address",
    "none",
    "n/a",
    "na",
    "unknown",
}

UNIT_ONLY_RE = re.compile(
    r"^\s*(?:apt|apartment|unit|suite|ste|space|lot|room|rm|#)\.?\s*[-#\w]+\s*$",
    re.IGNORECASE,
)

PO_BOX_RE = re.compile(
    r"\bP\.?\s*O\.?\s*Box\b|\bPOB\b|\bPMB\b",
    re.IGNORECASE,
)

# A likely physical street address begins with a building number.
STREET_START_RE = re.compile(r"^\s*\d+[A-Za-z]?(?:-\d+)?\s+\S+")

# Split a cell when a second numbered street address follows a long whitespace
# gap. This handles records such as:
# "PO Box 36 ...  25453 SW Forest Service Rd, Camp Sherman, OR 97730"
COMPOUND_ADDRESS_RE = re.compile(r"\s{2,}(?=\d+[A-Za-z]?(?:-\d+)?\s+\S+)")

ZIP_RE = re.compile(r"\b(\d{5})(?:-\d{4})?\b")
STATE_RE = re.compile(r"\b(OR|Oregon)\b", re.IGNORECASE)


class PassengerGeocodingError(RuntimeError):
    """Raised when passenger address preparation or geocoding fails."""


@dataclass
class AddressRecord:
    request_id: str
    passenger_name: str
    profile_id: str
    mobile_phone: str
    home_phone: str
    phone_numbers: str
    address_source: str
    input_address: str
    street: str
    city: str
    state: str
    zip_code: str
    geocode_status: str = "PENDING"
    match_type: str = ""
    matched_address: str = ""
    longitude: str = ""
    latitude: str = ""
    tigerline_id: str = ""
    tigerline_side: str = ""
    evacuation_zone: str = ""
    proximity_to_zone_miles: str = ""
    proximity_status: str = ""
    review_notes: str = ""

    def to_output_row(self) -> dict[str, str]:
        return {
            "passenger_name": self.passenger_name,
            "profile_id": self.profile_id,
            "mobile_phone": self.mobile_phone,
            "home_phone": self.home_phone,
            "phone_numbers": self.phone_numbers,
            "address_source": self.address_source,
            "input_address": self.input_address,
            "street": self.street,
            "city": self.city,
            "state": self.state,
            "zip": self.zip_code,
            "geocode_status": self.geocode_status,
            "match_type": self.match_type,
            "matched_address": self.matched_address,
            "longitude": self.longitude,
            "latitude": self.latitude,
            "tigerline_id": self.tigerline_id,
            "tigerline_side": self.tigerline_side,
            "evacuation_zone": self.evacuation_zone,
            "proximity_to_zone_miles": self.proximity_to_zone_miles,
            "proximity_status": self.proximity_status,
            "review_notes": self.review_notes,
        }


def clean_text(value: str | None) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", value).strip(" ,")


def full_name(firstname: str, lastname: str) -> str:
    return clean_text(f"{clean_text(firstname)} {clean_text(lastname)}")


def split_compound_address(value: str) -> list[str]:
    """
    Conservatively split cells containing two distinct addresses.

    Ordinary comma-separated address components are not split.
    """
    value = clean_text(value)
    if not value:
        return []

    parts = [clean_text(part) for part in COMPOUND_ADDRESS_RE.split(value)]
    return [part for part in parts if part]


def is_invalid_address(value: str) -> bool:
    normalized = clean_text(value).lower()
    return not normalized or normalized in INVALID_ADDRESS_PHRASES


def is_unit_only(value: str) -> bool:
    return bool(UNIT_ONLY_RE.fullmatch(clean_text(value)))


def is_po_box(value: str) -> bool:
    return bool(PO_BOX_RE.search(value))


def looks_like_street_address(value: str) -> bool:
    return bool(STREET_START_RE.match(clean_text(value)))


def parse_address_for_census(
    address: str,
    *,
    fallback_city: str = "",
    fallback_state: str = "OR",
) -> tuple[str, str, str, str]:
    """
    Parse a mostly comma-delimited U.S. address into Census batch fields.

    The original address remains in input_address for audit and manual review.
    """
    original = clean_text(address)
    working = re.sub(r",?\s*USA\s*$", "", original, flags=re.IGNORECASE)
    working = re.sub(
        r",?\s*United States\s*$",
        "",
        working,
        flags=re.IGNORECASE,
    )

    zip_match = ZIP_RE.search(working)
    zip_code = zip_match.group(1) if zip_match else ""

    state = "OR" if STATE_RE.search(working) else fallback_state

    parts = [clean_text(part) for part in working.split(",") if clean_text(part)]

    street = parts[0] if parts else working
    city = fallback_city

    if len(parts) >= 2:
        # Usually: street, city, state ZIP
        city_candidate = parts[-2]
        if not STATE_RE.fullmatch(city_candidate):
            city = city_candidate

    # Handle uncommaed examples such as:
    # "160 S Oak Street #235 Sisters OR 97759"
    if len(parts) == 1:
        match = re.match(
            r"^(.*?)(?:\s+)([A-Za-z][A-Za-z .'-]+?)\s+"
            r"(?:OR|Oregon)\s+\d{5}(?:-\d{4})?\s*$",
            working,
            flags=re.IGNORECASE,
        )
        if match:
            street = clean_text(match.group(1))
            city = clean_text(match.group(2))

    street = re.sub(
        r",?\s+[A-Za-z][A-Za-z .'-]+,\s*(?:OR|Oregon)\s+\d{5}.*$",
        "",
        street,
        flags=re.IGNORECASE,
    ).strip(" ,")

    return street, city, state, zip_code


def _new_request_id() -> str:
    # Random temporary ID prevents passenger profile IDs from being sent.
    return secrets.token_hex(8)


def expand_passenger_addresses(
    input_csv: str | Path,
) -> list[AddressRecord]:
    """
    Produce one AddressRecord per distinct passenger address.

    Rules:
    - address 1 is treated as the primary location.
    - A unit-only address 2 is appended to address 1.
    - A full address in address 2 becomes an additional record.
    - PO boxes remain in the report but are not sent for geocoding because
      they do not identify a physical evacuation location.
    - Clearly invalid notes remain in the report for manual review.
    """
    input_path = Path(input_csv)

    try:
        with input_path.open(
            "r",
            encoding="utf-8-sig",
            newline="",
        ) as file:
            reader = csv.DictReader(file)
            fieldnames = set(reader.fieldnames or [])
            missing = REQUIRED_COLUMNS - fieldnames
            if missing:
                raise PassengerGeocodingError(
                    "Input CSV is missing required columns: "
                    + ", ".join(sorted(missing))
                )
            source_rows = list(reader)
    except OSError as exc:
        raise PassengerGeocodingError(
            f"Could not read {input_path}: {exc}"
        ) from exc

    records: list[AddressRecord] = []

    for source in source_rows:
        passenger_name = full_name(
            source.get("firstname", ""),
            source.get("lastname", ""),
        )
        profile_id = clean_text(source.get("profile_id", ""))
        mobile_phone = clean_text(source.get("mobilephone_number", ""))
        home_phone = clean_text(source.get("homephone_number", ""))
        phone_numbers = " | ".join(
            number
            for number in dict.fromkeys([mobile_phone, home_phone])
            if number
        )
        address1 = clean_text(source.get("address 1", ""))
        address2 = clean_text(source.get("address 2", ""))

        candidate_addresses: list[tuple[str, str]] = []

        address2_parts = split_compound_address(address2)

        if (
            address1
            and len(address2_parts) == 1
            and is_unit_only(address2_parts[0])
        ):
            address1 = clean_text(f"{address1}, {address2_parts[0]}")
            address2_parts = []

        for index, part in enumerate(split_compound_address(address1), start=1):
            source_name = "address 1"
            if index > 1:
                source_name = f"address 1 part {index}"
            candidate_addresses.append((source_name, part))

        for index, part in enumerate(address2_parts, start=1):
            source_name = "address 2"
            if len(address2_parts) > 1:
                source_name = f"address 2 part {index}"
            candidate_addresses.append((source_name, part))

        if not candidate_addresses:
            records.append(
                AddressRecord(
                    request_id=_new_request_id(),
                    passenger_name=passenger_name,
                    profile_id=profile_id,
                    mobile_phone=mobile_phone,
                    home_phone=home_phone,
                    phone_numbers=phone_numbers,
                    address_source="",
                    input_address="",
                    street="",
                    city="",
                    state="",
                    zip_code="",
                    geocode_status="NO_ADDRESS",
                    review_notes="No address was present in either address field.",
                )
            )
            continue

        seen: set[str] = set()

        for source_name, address in candidate_addresses:
            duplicate_key = re.sub(r"\W+", "", address).lower()
            if duplicate_key in seen:
                continue
            seen.add(duplicate_key)

            street, city, state, zip_code = parse_address_for_census(address)

            record = AddressRecord(
                request_id=_new_request_id(),
                passenger_name=passenger_name,
                profile_id=profile_id,
                mobile_phone=mobile_phone,
                home_phone=home_phone,
                phone_numbers=phone_numbers,
                address_source=source_name,
                input_address=address,
                street=street,
                city=city,
                state=state,
                zip_code=zip_code,
            )

            if is_invalid_address(address):
                record.geocode_status = "INVALID_ADDRESS"
                record.review_notes = (
                    "Address field contains a note or unusable value."
                )
            elif is_po_box(address) and not looks_like_street_address(address):
                record.geocode_status = "MAILING_ONLY"
                record.review_notes = (
                    "PO Box or private mailbox does not identify a physical "
                    "evacuation location."
                )
            elif not looks_like_street_address(street):
                record.geocode_status = "REVIEW"
                record.review_notes = (
                    "Could not confidently identify a numbered street address."
                )
            elif not city and not zip_code:
                record.geocode_status = "REVIEW"
                record.review_notes = (
                    "Street address lacks both a city and ZIP code."
                )

            records.append(record)

    return records


def geocode_with_census(
    records: list[AddressRecord],
    *,
    benchmark: str = DEFAULT_BENCHMARK,
    timeout: int = DEFAULT_TIMEOUT_SECONDS,
    session: requests.Session | None = None,
) -> None:
    """
    Geocode eligible records in place with the Census batch endpoint.

    Only request_id and address components are transmitted. Passenger names
    and source profile IDs are not sent.
    """
    eligible = [
        record
        for record in records
        if record.geocode_status == "PENDING"
    ]

    if not eligible:
        LOGGER.info("No eligible physical addresses require geocoding.")
        return

    request_buffer = io.StringIO(newline="")
    writer = csv.writer(request_buffer, lineterminator="\n")

    for record in eligible:
        writer.writerow(
            [
                record.request_id,
                record.street,
                record.city,
                record.state,
                record.zip_code,
            ]
        )

    request_bytes = request_buffer.getvalue().encode("utf-8")

    owns_session = session is None
    http = session or requests.Session()
    http.headers.setdefault(
        "User-Agent",
        "PassengerEvacuationGeocoder/1.0",
    )

    try:
        response = http.post(
            CENSUS_BATCH_URL,
            data={"benchmark": benchmark},
            files={
                "addressFile": (
                    "addresses.csv",
                    request_bytes,
                    "text/csv",
                )
            },
            timeout=timeout,
        )
        response.raise_for_status()
    except requests.RequestException as exc:
        raise PassengerGeocodingError(
            f"Census geocoding request failed: {exc}"
        ) from exc
    finally:
        if owns_session:
            http.close()

    by_request_id = {record.request_id: record for record in eligible}

    decoded = response.content.decode("utf-8-sig", errors="replace")
    result_reader = csv.reader(io.StringIO(decoded))

    returned_ids: set[str] = set()

    for row in result_reader:
        if not row:
            continue

        # Census locations/addressbatch output:
        # 0 ID
        # 1 Input address
        # 2 Match status
        # 3 Match type
        # 4 Matched address
        # 5 Coordinates: "longitude,latitude"
        # 6 TIGER/Line ID
        # 7 Side
        row += [""] * (8 - len(row))

        request_id = row[0].strip()
        record = by_request_id.get(request_id)
        if record is None:
            LOGGER.warning(
                "Ignoring Census response with unknown request ID %s",
                request_id,
            )
            continue

        returned_ids.add(request_id)
        match_status = row[2].strip()
        record.match_type = row[3].strip()
        record.matched_address = row[4].strip()
        record.tigerline_id = row[6].strip()
        record.tigerline_side = row[7].strip()

        coordinates = row[5].strip()
        if "," in coordinates:
            longitude, latitude = coordinates.split(",", 1)
            record.longitude = longitude.strip()
            record.latitude = latitude.strip()

        if match_status.lower() == "match" and record.longitude and record.latitude:
            record.geocode_status = "MATCH"
        else:
            record.geocode_status = "NO_MATCH"
            record.review_notes = (
                "Census Geocoder did not return a usable coordinate match."
            )

    for record in eligible:
        if record.request_id not in returned_ids:
            record.geocode_status = "NO_RESPONSE"
            record.review_notes = (
                "No result row was returned by the Census Geocoder."
            )


def write_report(
    records: Iterable[AddressRecord],
    output_csv: str | Path,
) -> Path:
    output_path = Path(output_csv)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    temporary = output_path.with_suffix(output_path.suffix + ".tmp")

    try:
        with temporary.open(
            "w",
            encoding="utf-8-sig",
            newline="",
        ) as file:
            writer = csv.DictWriter(file, fieldnames=OUTPUT_COLUMNS)
            writer.writeheader()
            for record in records:
                writer.writerow(record.to_output_row())

        temporary.replace(output_path)
    except OSError as exc:
        try:
            temporary.unlink(missing_ok=True)
        except OSError:
            pass
        raise PassengerGeocodingError(
            f"Could not write report {output_path}: {exc}"
        ) from exc

    return output_path


def create_passenger_geocode_report(
    input_csv: str | Path,
    output_csv: str | Path,
    *,
    prepare_only: bool = False,
    benchmark: str = DEFAULT_BENCHMARK,
    timeout: int = DEFAULT_TIMEOUT_SECONDS,
) -> list[AddressRecord]:
    records = expand_passenger_addresses(input_csv)

    if not prepare_only:
        geocode_with_census(
            records,
            benchmark=benchmark,
            timeout=timeout,
        )

    write_report(records, output_csv)
    return records


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Expand and geocode passenger addresses for evacuation-zone "
            "analysis."
        )
    )
    parser.add_argument(
        "input_csv",
        type=Path,
        help="Passenger CSV exported from SPEDSTA.",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=Path("passenger_evacuation_report.csv"),
        help="Output report CSV.",
    )
    parser.add_argument(
        "--prepare-only",
        action="store_true",
        help=(
            "Expand and classify addresses without contacting the Census "
            "Geocoder."
        ),
    )
    parser.add_argument(
        "--benchmark",
        default=DEFAULT_BENCHMARK,
        help="Census Geocoder benchmark name.",
    )
    parser.add_argument(
        "--timeout",
        type=int,
        default=DEFAULT_TIMEOUT_SECONDS,
        help="HTTP timeout in seconds.",
    )
    parser.add_argument(
        "--quiet",
        action="store_true",
        help="Suppress informational logging.",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()

    logging.basicConfig(
        level=logging.WARNING if args.quiet else logging.INFO,
        format="%(levelname)s: %(message)s",
    )

    try:
        records = create_passenger_geocode_report(
            args.input_csv,
            args.output,
            prepare_only=args.prepare_only,
            benchmark=args.benchmark,
            timeout=args.timeout,
        )
    except PassengerGeocodingError as exc:
        LOGGER.error("%s", exc)
        return 1

    counts: dict[str, int] = {}
    for record in records:
        counts[record.geocode_status] = (
            counts.get(record.geocode_status, 0) + 1
        )

    print(f"Wrote {len(records)} address records to {args.output.resolve()}")
    for status, count in sorted(counts.items()):
        print(f"  {status}: {count}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())