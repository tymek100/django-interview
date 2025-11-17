import io
from typing import Any, Dict, List, Optional

from openpyxl import load_workbook
from rest_framework import parsers, status
from rest_framework.response import Response
from rest_framework.views import APIView

from .serializers import ExcelSummaryRequestSerializer, ExcelSummaryResponseSerializer


def normalize_header(header: Any) -> str:
    """Normalize header cell to allow case-insensitive matching."""
    if header is None:
        return ""
    return str(header).strip().lower()


def coerce_to_number(value: Any) -> Optional[float]:
    """
    Try to convert various cell values to float.

    Handles:
    - numeric types
    - strings like "$90,00", "90,00", "90.00", "1,234.56"
    """
    if value is None:
        return None

    if isinstance(value, (int, float)):
        return float(value)

    if isinstance(value, str):
        s = value.strip()
        if not s:
            return None

        # Remove common currency symbols and spaces
        for ch in ["$", "€", "£"]:
            s = s.replace(ch, "")
        s = s.replace(" ", "")

        # Handle decimal/thousands separators
        if "," in s and "." in s:
            # Assume '.' is decimal, ',' is thousands separator: "1,234.56"
            s = s.replace(",", "")
        elif "," in s and "." not in s:
            # European style: "90,00" -> "90.00"
            s = s.replace(",", ".")

        try:
            return float(s)
        except ValueError:
            return None

    return None


class ExcelSummaryView(APIView):
    """
    POST /api/excel-summary/

    multipart/form-data:
        file: <Excel file>
        columns: <column name>   (can appear multiple times)
        columns: <another column>

    Response:
    {
        "file": "<filename>",
        "summary": [
            {"column": "CURRENT USD", "sum": 1234.5, "avg": 56.7},
            ...
        ]
    }
    """

    parser_classes = [parsers.MultiPartParser, parsers.FormParser]

    def post(self, request, *args, **kwargs):
        request_serializer = ExcelSummaryRequestSerializer(data=request.data)
        request_serializer.is_valid(raise_exception=True)

        uploaded_file = request_serializer.validated_data["file"]
        columns: List[str] = request_serializer.validated_data["columns"]

        # Read workbook from in-memory bytes
        try:
            file_bytes = io.BytesIO(uploaded_file.read())
            wb = load_workbook(filename=file_bytes, data_only=True)
        except Exception:
            return Response(
                {"detail": "Unable to read Excel file. Make sure it is a valid .xlsx file."},
                status=status.HTTP_400_BAD_REQUEST,
            )

        ws = wb.active  # use first sheet

        # Find header row: first non-empty row
        header_row_index = None
        header_row_values = None
        for idx, row in enumerate(ws.iter_rows(min_row=1, max_row=20, values_only=True), start=1):
            if any(cell not in (None, "") for cell in row):
                header_row_index = idx
                header_row_values = row
                break

        if header_row_index is None or header_row_values is None:
            return Response(
                {"detail": "Could not detect header row in the Excel sheet."},
                status=status.HTTP_400_BAD_REQUEST,
            )

        # Map normalized header -> column index
        header_map: Dict[str, int] = {}
        for col_idx, cell in enumerate(header_row_values):
            key = normalize_header(cell)
            if key:
                header_map[key] = col_idx

        requested_summaries = []
        missing_columns = []

        # Iterate over requested columns and compute sum + avg
        for col_name in columns:
            norm = normalize_header(col_name)
            if norm not in header_map:
                missing_columns.append(col_name)
                continue

            col_idx = header_map[norm]
            total = 0.0
            count = 0

            # Iterate data rows, starting after header
            for row in ws.iter_rows(
                min_row=header_row_index + 1,
                max_row=ws.max_row,
                values_only=True,
            ):
                if col_idx >= len(row):
                    continue
                cell_value = row[col_idx]
                number = coerce_to_number(cell_value)
                if number is None:
                    continue
                total += number
                count += 1

            if count > 0:
                avg = total / count
                requested_summaries.append(
                    {
                        "column": col_name,
                        "sum": round(total, 2),
                        "avg": round(avg, 2),
                    }
                )
            else:
                # Column exists but no numeric values
                requested_summaries.append(
                    {
                        "column": col_name,
                        "sum": 0.0,
                        "avg": 0.0,
                    }
                )

        if not requested_summaries:
            return Response(
                {
                    "detail": "None of the requested columns were found in the sheet.",
                    "requested_columns": columns,
                    "available_columns": list(header_map.keys()),
                },
                status=status.HTTP_400_BAD_REQUEST,
            )

        response_data = {
            "file": uploaded_file.name,
            "summary": requested_summaries,
        }

        # Optional: validate response format with a serializer (nice for docs / schema)
        response_serializer = ExcelSummaryResponseSerializer(data=response_data)
        response_serializer.is_valid(raise_exception=True)

        return Response(response_serializer.data, status=status.HTTP_200_OK)
