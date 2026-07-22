"""
oregon_evacuation_zones.py

Download Oregon's current public evacuation-area polygons as GeoJSON.

Source:
https://oregon-oem-geo.hub.arcgis.com/datasets/evacuation-areas-public

Dependency:
    pip install requests

Examples:
    python oregon_evacuation_zones.py
    python oregon_evacuation_zones.py --output data/oregon_evacuation_zones.geojson

Import usage:
    from oregon_evacuation_zones import download_oregon_evacuation_zones

    geojson = download_oregon_evacuation_zones(
        output_path="oregon_evacuation_zones.geojson"
    )
"""

from __future__ import annotations

import argparse
import json
import logging
from pathlib import Path
from typing import Any
from urllib.parse import urlparse

import requests


LOGGER = logging.getLogger(__name__)

OREGON_EVACUATION_ITEM_ID = "dd2bee51e3004dc4b1004df34fbd9e29"

ARCGIS_ITEM_URL = (
    "https://www.arcgis.com/sharing/rest/content/items/"
    f"{OREGON_EVACUATION_ITEM_ID}"
)

DEFAULT_OUTPUT_PATH = Path("data/oregon_evacuation_zones.geojson")
DEFAULT_TIMEOUT_SECONDS = 30
DEFAULT_BATCH_SIZE = 200


class EvacuationDownloadError(RuntimeError):
    """Raised when the Oregon evacuation data cannot be downloaded."""


def _request_json(
    session: requests.Session,
    url: str,
    *,
    method: str = "GET",
    params: dict[str, Any] | None = None,
    data: dict[str, Any] | None = None,
    timeout: int = DEFAULT_TIMEOUT_SECONDS,
) -> dict[str, Any]:
    """Perform an HTTP request and return a validated JSON object."""

    try:
        if method.upper() == "POST":
            response = session.post(
                url,
                params=params,
                data=data,
                timeout=timeout,
            )
        else:
            response = session.get(
                url,
                params=params,
                timeout=timeout,
            )

        response.raise_for_status()

    except requests.RequestException as exc:
        raise EvacuationDownloadError(
            f"{method.upper()} request failed for {url}: {exc}"
        ) from exc

    try:
        result = response.json()
    except requests.exceptions.JSONDecodeError as exc:
        raise EvacuationDownloadError(
            f"ArcGIS returned invalid JSON for {response.url}"
        ) from exc

    if not isinstance(result, dict):
        raise EvacuationDownloadError(
            f"Expected a JSON object from {response.url}"
        )

    if "error" in result:
        error = result["error"]
        message = error.get("message", "Unknown ArcGIS error")
        details = error.get("details", [])

        if details:
            message = f"{message}: {'; '.join(details)}"

        raise EvacuationDownloadError(
            f"ArcGIS error from {response.url}: {message}"
        )

    return result


def resolve_feature_service_url(
    session: requests.Session,
    *,
    timeout: int = DEFAULT_TIMEOUT_SECONDS,
) -> str:
    """
    Resolve the ArcGIS item to its current FeatureServer URL.

    This avoids relying on a service URL that might later be republished
    or changed while the public ArcGIS item remains the same.
    """

    item = _request_json(
        session,
        ARCGIS_ITEM_URL,
        params={"f": "json"},
        timeout=timeout,
    )

    service_url = item.get("url")

    if not isinstance(service_url, str) or not service_url.strip():
        raise EvacuationDownloadError(
            "The Oregon evacuation ArcGIS item did not contain a service URL."
        )

    service_url = service_url.rstrip("/")

    if "FeatureServer" not in service_url:
        raise EvacuationDownloadError(
            f"Expected a FeatureServer URL, received: {service_url}"
        )

    return service_url


def _url_has_layer_number(url: str) -> bool:
    """Return True when an ArcGIS URL ends with a numeric layer ID."""

    path = urlparse(url).path.rstrip("/")
    final_component = path.rsplit("/", 1)[-1]
    return final_component.isdigit()


def find_polygon_layer_url(
    session: requests.Session,
    service_url: str,
    *,
    timeout: int = DEFAULT_TIMEOUT_SECONDS,
) -> str:
    """
    Locate the polygon layer within the FeatureServer.

    If the item already points directly to a numbered layer, it is
    validated and returned.
    """

    if _url_has_layer_number(service_url):
        layer = _request_json(
            session,
            service_url,
            params={"f": "json"},
            timeout=timeout,
        )

        geometry_type = layer.get("geometryType")
        if geometry_type != "esriGeometryPolygon":
            raise EvacuationDownloadError(
                "The ArcGIS item points to a layer that is not polygon data: "
                f"{geometry_type!r}"
            )

        return service_url

    service = _request_json(
        session,
        service_url,
        params={"f": "json"},
        timeout=timeout,
    )

    layers = service.get("layers", [])

    if not isinstance(layers, list) or not layers:
        raise EvacuationDownloadError(
            f"No layers were found in FeatureServer: {service_url}"
        )

    polygon_candidates: list[tuple[int, str]] = []

    for layer_summary in layers:
        layer_id = layer_summary.get("id")
        layer_name = str(layer_summary.get("name", ""))

        if not isinstance(layer_id, int):
            continue

        layer_url = f"{service_url}/{layer_id}"
        layer = _request_json(
            session,
            layer_url,
            params={"f": "json"},
            timeout=timeout,
        )

        if layer.get("geometryType") == "esriGeometryPolygon":
            polygon_candidates.append((layer_id, layer_name))

    if not polygon_candidates:
        raise EvacuationDownloadError(
            "No polygon layers were found in the evacuation FeatureServer."
        )

    if len(polygon_candidates) > 1:
        LOGGER.warning(
            "Multiple polygon layers were found. Using layer %s (%s).",
            polygon_candidates[0][0],
            polygon_candidates[0][1],
        )

    selected_layer_id, selected_layer_name = polygon_candidates[0]

    LOGGER.info(
        "Selected polygon layer %s: %s",
        selected_layer_id,
        selected_layer_name,
    )

    return f"{service_url}/{selected_layer_id}"


def get_object_ids(
    session: requests.Session,
    layer_url: str,
    *,
    where: str = "1=1",
    timeout: int = DEFAULT_TIMEOUT_SECONDS,
) -> list[int]:
    """Request the IDs of every feature matching the query."""

    result = _request_json(
        session,
        f"{layer_url}/query",
        params={
            "where": where,
            "returnIdsOnly": "true",
            "f": "json",
        },
        timeout=timeout,
    )

    object_ids = result.get("objectIds")

    # ArcGIS may return null when no records exist.
    if object_ids is None:
        return []

    if not isinstance(object_ids, list):
        raise EvacuationDownloadError(
            "ArcGIS returned an invalid objectIds response."
        )

    valid_ids = [value for value in object_ids if isinstance(value, int)]
    valid_ids.sort()

    return valid_ids


def _download_feature_batch(
    session: requests.Session,
    layer_url: str,
    object_ids: list[int],
    *,
    out_fields: str = "*",
    timeout: int = DEFAULT_TIMEOUT_SECONDS,
) -> dict[str, Any]:
    """Download one batch of features as GeoJSON using HTTP POST."""

    result = _request_json(
        session,
        f"{layer_url}/query",
        method="POST",
        data={
            "objectIds": ",".join(str(value) for value in object_ids),
            "outFields": out_fields,
            "returnGeometry": "true",
            "outSR": "4326",
            "f": "geojson",
        },
        timeout=timeout,
    )

    if result.get("type") != "FeatureCollection":
        raise EvacuationDownloadError(
            "ArcGIS did not return a GeoJSON FeatureCollection."
        )

    features = result.get("features")

    if not isinstance(features, list):
        raise EvacuationDownloadError(
            "The GeoJSON response did not contain a valid features array."
        )

    return result


def fetch_oregon_evacuation_zones(
    *,
    where: str = "1=1",
    out_fields: str = "*",
    batch_size: int = DEFAULT_BATCH_SIZE,
    timeout: int = DEFAULT_TIMEOUT_SECONDS,
    session: requests.Session | None = None,
) -> dict[str, Any]:
    """
    Fetch all current Oregon evacuation polygons.

    Returns:
        A GeoJSON FeatureCollection dictionary.

    Notes:
        An empty FeatureCollection is valid and ordinarily means that no
        evacuation polygons currently match the supplied query.
    """

    if batch_size < 1:
        raise ValueError("batch_size must be at least 1")

    owns_session = session is None
    http = session or requests.Session()

    http.headers.setdefault(
        "User-Agent",
        "OregonEvacuationZoneDownloader/1.0",
    )

    try:
        service_url = resolve_feature_service_url(
            http,
            timeout=timeout,
        )

        LOGGER.info("Resolved FeatureServer: %s", service_url)

        layer_url = find_polygon_layer_url(
            http,
            service_url,
            timeout=timeout,
        )

        LOGGER.info("Using evacuation layer: %s", layer_url)

        object_ids = get_object_ids(
            http,
            layer_url,
            where=where,
            timeout=timeout,
        )

        LOGGER.info("Found %d evacuation polygons.", len(object_ids))

        if not object_ids:
            return {
                "type": "FeatureCollection",
                "features": [],
            }

        all_features: list[dict[str, Any]] = []

        for start in range(0, len(object_ids), batch_size):
            batch_ids = object_ids[start : start + batch_size]

            LOGGER.info(
                "Downloading features %d through %d of %d.",
                start + 1,
                start + len(batch_ids),
                len(object_ids),
            )

            batch = _download_feature_batch(
                http,
                layer_url,
                batch_ids,
                out_fields=out_fields,
                timeout=timeout,
            )

            all_features.extend(batch["features"])

        return {
            "type": "FeatureCollection",
            "features": all_features,
        }

    finally:
        if owns_session:
            http.close()


def save_geojson(
    geojson: dict[str, Any],
    output_path: str | Path,
) -> Path:
    """Write a GeoJSON object to disk atomically."""

    path = Path(output_path)
    path.parent.mkdir(parents=True, exist_ok=True)

    temporary_path = path.with_name(f"{path.name}.tmp")

    try:
        temporary_path.write_text(
            json.dumps(
                geojson,
                ensure_ascii=False,
                indent=2,
            )
            + "\n",
            encoding="utf-8",
        )

        temporary_path.replace(path)
    except OSError as exc:
        try:
            temporary_path.unlink(missing_ok=True)
        except OSError:
            pass

        raise EvacuationDownloadError(
            f"Could not save GeoJSON to {path}: {exc}"
        ) from exc

    return path


def download_oregon_evacuation_zones(
    output_path: str | Path = DEFAULT_OUTPUT_PATH,
    *,
    where: str = "1=1",
    out_fields: str = "*",
    batch_size: int = DEFAULT_BATCH_SIZE,
    timeout: int = DEFAULT_TIMEOUT_SECONDS,
) -> dict[str, Any]:
    """
    Download all matching evacuation polygons and save them to disk.

    Returns:
        The downloaded GeoJSON FeatureCollection.
    """

    geojson = fetch_oregon_evacuation_zones(
        where=where,
        out_fields=out_fields,
        batch_size=batch_size,
        timeout=timeout,
    )

    saved_path = save_geojson(geojson, output_path)

    LOGGER.info(
        "Saved %d features to %s.",
        len(geojson["features"]),
        saved_path.resolve(),
    )

    return geojson


def _parse_arguments() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Download Oregon's current public evacuation-area polygons "
            "as GeoJSON."
        )
    )

    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=DEFAULT_OUTPUT_PATH,
        help=(
            "Output GeoJSON filename. "
            f"Default: {DEFAULT_OUTPUT_PATH}"
        ),
    )

    parser.add_argument(
        "--where",
        default="1=1",
        help=(
            "Optional ArcGIS SQL where clause. "
            "Default: 1=1"
        ),
    )

    parser.add_argument(
        "--batch-size",
        type=int,
        default=DEFAULT_BATCH_SIZE,
        help=(
            "Maximum features requested per ArcGIS query. "
            f"Default: {DEFAULT_BATCH_SIZE}"
        ),
    )

    parser.add_argument(
        "--timeout",
        type=int,
        default=DEFAULT_TIMEOUT_SECONDS,
        help=(
            "HTTP timeout in seconds. "
            f"Default: {DEFAULT_TIMEOUT_SECONDS}"
        ),
    )

    parser.add_argument(
        "--quiet",
        action="store_true",
        help="Suppress informational logging.",
    )

    return parser.parse_args()


def main() -> int:
    args = _parse_arguments()

    logging.basicConfig(
        level=logging.WARNING if args.quiet else logging.INFO,
        format="%(levelname)s: %(message)s",
    )

    try:
        geojson = download_oregon_evacuation_zones(
            output_path=args.output,
            where=args.where,
            batch_size=args.batch_size,
            timeout=args.timeout,
        )
    except (EvacuationDownloadError, ValueError) as exc:
        LOGGER.error("%s", exc)
        return 1

    print(
        f"Downloaded {len(geojson['features'])} evacuation polygons "
        f"to {args.output.resolve()}"
    )

    return 0


if __name__ == "__main__":
    raise SystemExit(main())