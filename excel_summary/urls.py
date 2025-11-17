from django.urls import path

from .views import ExcelSummaryView

urlpatterns = [
    path("excel-summary/", ExcelSummaryView.as_view(), name="excel-summary"),
]
