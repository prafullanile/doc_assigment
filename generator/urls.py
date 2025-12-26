from django.urls import path
from .views import generate_doc

from django.urls import path
from .views import generate_doc

urlpatterns = [
    path('', generate_doc, name='home'),
    path('generate/', generate_doc, name='generate_doc'),
]
