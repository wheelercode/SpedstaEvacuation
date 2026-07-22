"""
geocode_passengers.py

Expand passenger address fields into one record per usable address, geocode
those addresses with Oregon address points and the U.S. Census Geocoder, and
create a report-ready CSV for later evacuation-zone analysis.

The module does NOT determine evacuation zones. It leaves evacuation fields
blank so a later spatial-analysis module can populate them.

Input expected:
    data/spedsta_passengers.csv

Required columns:
    profile_id
    firstname
    lastname
    address 1
    address 2

Dependency:
    pip install requests

Examples:
    python geocode_passengers.py data/spedsta_passengers.csv
    python geocode_passengers.py data/spedsta_passengers.csv -o passenger_evacuation_report.csv
    python geocode_passengers.py data/spedsta_passengers.csv --prepare-only

Privacy note:
    This sends address text, but not passenger names, phone numbers, email
    addresses, or profile IDs, to the address geocoding services.
"""

from __future__ import annotations

import argparse
import csv
import logging
import re
import secrets
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

import requests


LOGGER = logging.getLogger(__name__)

DESCHUTES_E911_QUERY_URL = (
    "https://services1.arcgis.com/znO8Hz1SuVVohYhZ/"
    "ArcGIS/rest/services/E911_Address_Points/"
    "FeatureServer/0/query"
)
OREGON_ADDRESS_POINTS_QUERY_URL = (
    "https://services8.arcgis.com/8PAo5HGmvRMlF2eU/"
    "arcgis/rest/services/Oregon_Address_Points/"
    "FeatureServer/0/query"
)
CENSUS_ONE_LINE_URL = (
    "https://geocoding.geo.census.gov/"
    "geocoder/locations/onelineaddress"
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
    "geocode_source",
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

# Secondary-unit text is retained in input_address, but is not included in
# requests to address-point or Census services. This expression is applied to
# the parsed street component, not the original display address.
SECONDARY_UNIT_RE = re.compile(
    r"(?:\s*,?\s+)"
    r"(?:apt|apartment|unit|suite|ste|space|lot|room|rm|floor|fl|building|bldg|#)"
    r"\.?\s*[-#A-Za-z0-9]+(?:\s+[-#A-Za-z0-9]+)*\s*$",
    re.IGNORECASE,
)

STREET_COMPONENT_RE = re.compile(
    r"^\s*(?P<number>\d+)(?P<number_suffix>[A-Za-z]?)"
    r"(?:-\d+)?\s+(?P<street>.+?)\s*$"
)

STREET_TOKEN_ALIASES = {
    "lane": "ln",
    "ln": "ln",
    "road": "rd",
    "rd": "rd",
    "street": "st",
    "st": "st",
    "avenue": "ave",
    "ave": "ave",
    "drive": "dr",
    "dr": "dr",
    "highway": "hwy",
    "hwy": "hwy",
    "north": "n",
    "n": "n",
    "south": "s",
    "s": "s",
    "east": "e",
    "e": "e",
    "west": "w",
    "w": "w",
    "northeast": "ne",
    "ne": "ne",
    "northwest": "nw",
    "nw": "nw",
    "southeast": "se",
    "se": "se",
    "southwest": "sw",
    "sw": "sw",
}

DIRECTION_TOKENS = {"n", "s", "e", "w", "ne", "nw", "se", "sw"}
STREET_TYPE_TOKENS = {"ln", "rd", "st", "ave", "dr", "hwy"}

ARCGIS_COMMON_OUT_FIELDS = [
    "Add_Number",
    "AddNum_Suf",
    "St_PreMod",
    "St_PreDir",
    "St_PreTyp",
    "St_Name",
    "St_PosTyp",
    "St_PosDir",
    "St_PosMod",
    "Post_Comm",
    "Post_Code",
    "State",
    "Unit",
]

ARCGIS_LAYER_CONFIGS = [
    (
        DESCHUTES_E911_QUERY_URL,
        "DESCHUTES_E911",
        ["ADDRESS", "SUB_ADD_UNIT"],
    ),
    (
        OREGON_ADDRESS_POINTS_QUERY_URL,
        "OREGON_ADDRESS_POINTS",
        ["ADDRESS_FULL", "ADDRESS_NUMBER_FULL", "SUBADDRESS_FULL"],
    ),
]

ARCGIS_STREET_ATTRIBUTE_FIELDS = [
    "St_PreMod",
    "St_PreDir",
    "St_PreTyp",
    "St_Name",
    "St_PosTyp",
    "St_PosDir",
    "St_PosMod",
]


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
    geocode_source: str = ""
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
            "geocode_source": self.geocode_source,
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

    # Use the last five-digit component so a five-digit house number (for
    # example, 67667 Highway 20) does not mask the actual trailing ZIP code.
    zip_matches = list(ZIP_RE.finditer(working))
    zip_match = zip_matches[-1] if zip_matches else None
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


def strip_secondary_unit(street: str) -> str:
    """Return street text suitable for geocoding without changing display data."""
    return clean_text(SECONDARY_UNIT_RE.sub("", clean_text(street)))


def _normalize_address_tokens(value: object) -> tuple[str, ...]:
    tokens = re.findall(r"[A-Za-z0-9]+", clean_text(str(value or "")).lower())
    return tuple(STREET_TOKEN_ALIASES.get(token, token) for token in tokens)


def _street_request_parts(street: str) -> dict[str, object] | None:
    geocode_street = strip_secondary_unit(street)
    match = STREET_COMPONENT_RE.match(geocode_street)
    if not match:
        return None

    street_tokens = _normalize_address_tokens(match.group("street"))
    name_tokens = tuple(
        token
        for token in street_tokens
        if token not in DIRECTION_TOKENS and token not in STREET_TYPE_TOKENS
    )
    if not name_tokens:
        return None

    directions = tuple(
        token for token in street_tokens if token in DIRECTION_TOKENS
    )
    street_types = tuple(
        token for token in street_tokens if token in STREET_TYPE_TOKENS
    )

    return {
        "number": int(match.group("number")),
        "number_suffix": match.group("number_suffix").lower(),
        "street": geocode_street,
        "street_tokens": street_tokens,
        "name": " ".join(name_tokens),
        "directions": directions,
        "street_types": street_types,
    }


def _candidate_street(attributes: dict[str, object]) -> str:
    return clean_text(
        " ".join(
            str(attributes.get(field) or "")
            for field in ARCGIS_STREET_ATTRIBUTE_FIELDS
        )
    )


def _candidate_score(
    record: AddressRecord,
    requested: dict[str, object],
    attributes: dict[str, object],
) -> int | None:
    try:
        returned_number = int(attributes.get("Add_Number"))
    except (TypeError, ValueError):
        return None

    # A matching street name is mandatory; a house-number-only result is never
    # accepted. Number suffixes must also agree when the source supplies one.
    if returned_number != requested["number"]:
        return None

    returned_suffix = clean_text(
        str(attributes.get("AddNum_Suf") or "")
    ).lower()
    if returned_suffix != requested["number_suffix"]:
        return None

    returned_name = " ".join(
        _normalize_address_tokens(attributes.get("St_Name"))
    )
    if returned_name != requested["name"]:
        return None

    returned_street_tokens = _normalize_address_tokens(
        _candidate_street(attributes)
    )
    returned_directions = tuple(
        token
        for token in returned_street_tokens
        if token in DIRECTION_TOKENS
    )
    returned_types = tuple(
        token
        for token in returned_street_tokens
        if token in STREET_TYPE_TOKENS
    )

    requested_directions = requested["directions"]
    requested_types = requested["street_types"]
    if (
        requested_directions
        and returned_directions
        and requested_directions != returned_directions
    ):
        return None
    if requested_types and returned_types and requested_types != returned_types:
        return None

    score = 100
    if returned_street_tokens == requested["street_tokens"]:
        score += 50
    if requested_directions and requested_directions == returned_directions:
        score += 15
    if requested_types and requested_types == returned_types:
        score += 15

    returned_city = _normalize_address_tokens(attributes.get("Post_Comm"))
    requested_city = _normalize_address_tokens(record.city)
    if requested_city and returned_city:
        score += 25 if requested_city == returned_city else -25

    returned_zip = clean_text(str(attributes.get("Post_Code") or ""))[:5]
    if record.zip_code and returned_zip:
        score += 30 if record.zip_code == returned_zip else -30

    returned_state = clean_text(str(attributes.get("State") or "")).upper()
    if record.state and returned_state:
        score += 5 if record.state.upper() == returned_state else -20

    # Prefer the base address point after the input unit has been intentionally
    # removed, rather than selecting an arbitrary apartment/unit point.
    if not clean_text(str(attributes.get("Unit") or "")):
        score += 5

    return score


def _returned_arcgis_address(attributes: dict[str, object]) -> str:
    returned = clean_text(
        str(
            attributes.get("ADDRESS")
            or attributes.get("ADDRESS_FULL")
            or ""
        )
    )
    if returned:
        return returned

    number = clean_text(
        str(
            attributes.get("ADDRESS_NUMBER_FULL")
            or attributes.get("Add_Number")
            or ""
        )
    )
    parts = [number, _candidate_street(attributes)]
    city = clean_text(str(attributes.get("Post_Comm") or ""))
    state_zip = clean_text(
        f"{attributes.get('State') or ''} {attributes.get('Post_Code') or ''}"
    )
    if city:
        parts.append(city)
    if state_zip:
        parts.append(state_zip)
    return ", ".join(part for part in parts if part)


def _apply_match(
    record: AddressRecord,
    *,
    source: str,
    match_type: str,
    matched_address: str,
    longitude: object,
    latitude: object,
) -> bool:
    if longitude is None or latitude is None:
        return False

    try:
        longitude_value = float(longitude)
        latitude_value = float(latitude)
    except (TypeError, ValueError):
        return False

    record.geocode_status = "MATCH"
    record.geocode_source = source
    record.match_type = match_type
    record.matched_address = clean_text(matched_address)
    record.longitude = str(longitude_value)
    record.latitude = str(latitude_value)
    record.review_notes = ""
    return True


def _geocode_arcgis_record(
    record: AddressRecord,
    *,
    query_url: str,
    source: str,
    layer_fields: list[str],
    timeout: int,
    session: requests.Session,
) -> bool:
    requested = _street_request_parts(record.street)
    if requested is None:
        return False

    street_name = str(requested["name"]).replace("'", "''")
    response = session.get(
        query_url,
        params={
            "where": (
                f"Add_Number = {requested['number']} "
                f"AND St_Name = '{street_name}'"
            ),
            "outFields": ",".join(ARCGIS_COMMON_OUT_FIELDS + layer_fields),
            "returnGeometry": "true",
            "outSR": "4326",
            "f": "json",
        },
        timeout=timeout,
    )
    response.raise_for_status()
    payload = response.json()
    if not isinstance(payload, dict):
        raise PassengerGeocodingError("ArcGIS returned an invalid response.")
    if payload.get("error"):
        raise PassengerGeocodingError(
            f"ArcGIS response error: {payload['error']}"
        )

    candidates: list[tuple[int, dict[str, object], dict[str, object]]] = []
    for feature in payload.get("features") or []:
        attributes = feature.get("attributes") or {}
        geometry = feature.get("geometry") or {}
        score = _candidate_score(record, requested, attributes)
        if (
            score is not None
            and geometry.get("x") is not None
            and geometry.get("y") is not None
        ):
            candidates.append((score, attributes, geometry))

    if not candidates:
        return False

    _, attributes, geometry = max(candidates, key=lambda candidate: candidate[0])
    return _apply_match(
        record,
        source=source,
        match_type="ADDRESS_POINT",
        matched_address=_returned_arcgis_address(attributes),
        longitude=geometry.get("x"),
        latitude=geometry.get("y"),
    )


def _one_line_geocoding_address(record: AddressRecord) -> str:
    street = strip_secondary_unit(record.street)
    state_zip = clean_text(f"{record.state} {record.zip_code}")
    return ", ".join(
        part for part in [street, record.city, state_zip] if part
    )


def _geocode_census_record(
    record: AddressRecord,
    *,
    benchmark: str,
    timeout: int,
    session: requests.Session,
) -> bool:
    response = session.get(
        CENSUS_ONE_LINE_URL,
        params={
            "address": _one_line_geocoding_address(record),
            "benchmark": benchmark,
            "format": "json",
        },
        timeout=timeout,
    )
    response.raise_for_status()
    payload = response.json()
    if not isinstance(payload, dict):
        raise PassengerGeocodingError("Census returned an invalid response.")
    result = payload.get("result") or {}
    if not isinstance(result, dict):
        raise PassengerGeocodingError("Census returned an invalid result.")
    matches = result.get("addressMatches") or []
    if not matches:
        return False

    match = matches[0]
    coordinates = match.get("coordinates") or {}
    matched = _apply_match(
        record,
        source="CENSUS",
        match_type="CENSUS",
        matched_address=str(match.get("matchedAddress") or ""),
        longitude=coordinates.get("x"),
        latitude=coordinates.get("y"),
    )
    if matched:
        tiger_line = match.get("tigerLine") or {}
        record.tigerline_id = clean_text(
            str(tiger_line.get("tigerLineId") or "")
        )
        record.tigerline_side = clean_text(
            str(tiger_line.get("side") or "")
        )
    return matched


def geocode_addresses(
    records: list[AddressRecord],
    *,
    benchmark: str = DEFAULT_BENCHMARK,
    timeout: int = DEFAULT_TIMEOUT_SECONDS,
    session: requests.Session | None = None,
) -> None:
    """Geocode PENDING records in Deschutes, Oregon, Census order."""
    eligible = [
        record for record in records if record.geocode_status == "PENDING"
    ]
    if not eligible:
        LOGGER.info("No eligible physical addresses require geocoding.")
        return

    owns_session = session is None
    http = session or requests.Session()
    http.headers.setdefault("User-Agent", "PassengerEvacuationGeocoder/2.0")

    try:
        for record in eligible:
            matched = False
            for query_url, source, layer_fields in ARCGIS_LAYER_CONFIGS:
                try:
                    matched = _geocode_arcgis_record(
                        record,
                        query_url=query_url,
                        source=source,
                        layer_fields=layer_fields,
                        timeout=timeout,
                        session=http,
                    )
                except (requests.RequestException, PassengerGeocodingError, ValueError) as exc:
                    LOGGER.error(
                        "%s geocoding failed for address request %s: %s",
                        source,
                        record.request_id,
                        exc,
                    )
                if matched:
                    break

            if not matched:
                try:
                    matched = _geocode_census_record(
                        record,
                        benchmark=benchmark,
                        timeout=timeout,
                        session=http,
                    )
                except (requests.RequestException, PassengerGeocodingError, ValueError) as exc:
                    LOGGER.error(
                        "CENSUS geocoding failed for address request %s: %s",
                        record.request_id,
                        exc,
                    )

            if not matched:
                record.geocode_status = "NO_MATCH"
                record.geocode_source = ""
                record.match_type = ""
                record.matched_address = ""
                record.longitude = ""
                record.latitude = ""
                record.review_notes = (
                    "Deschutes E911, Oregon Address Points, and Census "
                    "did not return a usable coordinate match."
                )
    finally:
        if owns_session:
            http.close()


def geocode_with_census(
    records: list[AddressRecord],
    *,
    benchmark: str = DEFAULT_BENCHMARK,
    timeout: int = DEFAULT_TIMEOUT_SECONDS,
    session: requests.Session | None = None,
) -> None:
    """Backward-compatible Census-only wrapper for existing importers."""
    eligible = [
        record for record in records if record.geocode_status == "PENDING"
    ]
    if not eligible:
        LOGGER.info("No eligible physical addresses require geocoding.")
        return

    owns_session = session is None
    http = session or requests.Session()
    http.headers.setdefault("User-Agent", "PassengerEvacuationGeocoder/2.0")
    try:
        for record in eligible:
            try:
                matched = _geocode_census_record(
                    record,
                    benchmark=benchmark,
                    timeout=timeout,
                    session=http,
                )
            except (requests.RequestException, PassengerGeocodingError, ValueError) as exc:
                LOGGER.error(
                    "CENSUS geocoding failed for address request %s: %s",
                    record.request_id,
                    exc,
                )
                matched = False

            if not matched:
                record.geocode_status = "NO_MATCH"
                record.geocode_source = ""
                record.match_type = ""
                record.matched_address = ""
                record.longitude = ""
                record.latitude = ""
                record.review_notes = (
                    "Census Geocoder did not return a usable coordinate match."
                )
    finally:
        if owns_session:
            http.close()


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
        geocode_addresses(
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
        default=Path("data/passenger_address_coordinates.csv"),
        help="Output report CSV.",
    )
    parser.add_argument(
        "--prepare-only",
        action="store_true",
        help=(
            "Expand and classify addresses without contacting geocoding "
            "services."
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