from rest_framework import serializers


class ExcelSummaryRequestSerializer(serializers.Serializer):
    """
    Request payload for summarizing an uploaded Excel file.

    This is used as the request body schema for Swagger / OpenAPI.
    """
    file = serializers.FileField(
        help_text="Excel file in .xlsx format."
    )
    columns = serializers.CharField(
        default=["CURRENT USD", "CURRENT CAD"],
        help_text="List of column header names to summarize in Python-style list"
                  "e.g.: ['CURRENT USD', 'CURRENT CAD']",
    )


class ColumnSummarySerializer(serializers.Serializer):
    column = serializers.CharField()
    sum = serializers.FloatField()
    avg = serializers.FloatField()


class ExcelSummaryResponseSerializer(serializers.Serializer):
    """
    Response payload containing per-column sum and average.
    """
    file = serializers.CharField()
    summary = ColumnSummarySerializer(many=True)
