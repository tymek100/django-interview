import ast
import io

from openpyxl import load_workbook

from rest_framework import parsers, status
from rest_framework.response import Response
from rest_framework.views import APIView

from drf_spectacular.utils import (
    OpenApiResponse,
    extend_schema,
)

from .excel_utils import HeaderNotFoundError, summarize_excel_columns
from .serializers import (
    ExcelSummaryRequestSerializer,
    ExcelSummaryResponseSerializer,
)


class ExcelSummaryView(APIView):
    """
    POST /api/excel-summary/

    multipart/form-data:
        file: <Excel file>
        columns: <column name>   (can appear multiple times)

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

    @extend_schema(
        tags=["Excel Upload"],
        summary="Summarize numeric columns in an Excel file",
        description=(
            "Upload an Excel `.xlsx` file and provide a list of column header names.\n\n"
            "The API will locate each column and return the sum and average of "
            "all numeric values found in that column."
        ),
        request=ExcelSummaryRequestSerializer,
        responses={
            200: ExcelSummaryResponseSerializer,
            400: OpenApiResponse(
                description="Validation error or Excel parsing error."
            ),
        },
    )
    def post(self, request, *args, **kwargs):
        request_serializer = ExcelSummaryRequestSerializer(data=request.data)
        request_serializer.is_valid(raise_exception=True)

        uploaded_file = request_serializer.validated_data["file"]
        columns = ast.literal_eval(request_serializer.validated_data["columns"])

        try:
            file_bytes = io.BytesIO(uploaded_file.read())
            wb = load_workbook(filename=file_bytes, data_only=True)
        except Exception:
            return Response(
                {"detail": "Unable to read Excel file. Make sure it is a valid .xlsx file."},
                status=status.HTTP_400_BAD_REQUEST,
            )

        ws = wb.active  # use first sheet

        try:
            summaries, missing_columns, available_columns = summarize_excel_columns(ws, columns)
        except HeaderNotFoundError as exc:
            return Response(
                {"detail": str(exc)},
                status=status.HTTP_400_BAD_REQUEST,
            )

        if not summaries:
            return Response(
                {
                    "detail": "None of the requested columns were found in the sheet.",
                    "requested_columns": columns,
                    "available_columns": available_columns,
                },
                status=status.HTTP_400_BAD_REQUEST,
            )

        response_data = {
            "file": uploaded_file.name,
            "summary": summaries,
        }

        response_serializer = ExcelSummaryResponseSerializer(data=response_data)
        response_serializer.is_valid(raise_exception=True)

        return Response(response_serializer.data, status=status.HTTP_200_OK)
