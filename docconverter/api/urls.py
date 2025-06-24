from django.urls import path
from .views import DocConverterView

urlpatterns = [
    path('convert/', DocConverterView.as_view(), name='doc-converter'),
] 