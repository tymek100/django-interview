from rest_framework import serializers


class ExcelSummaryRequestSerializer(serializers.Serializer):
    """
    Request payload for summarizing an uploaded Excel file.
    - file: Excel file (.xlsx)
    - columns: list of column header names to summarize
    """
    file = serializers.FileField()
    columns = serializers.ListField(
        child=serializers.CharField(),
        allow_empty=False,
        help_text="List of column header names to summarize (e.g. ['CURRENT USD', 'CURRENT CAD'])",
    )


class ColumnSummarySerializer(serializers.Serializer):
    column = serializers.CharField()
    sum = serializers.FloatField()
    avg = serializers.FloatField()


class ExcelSummaryResponseSerializer(serializers.Serializer):
    file = serializers.CharField()
    summary = ColumnSummarySerializer(many=True)
