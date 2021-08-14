from django.urls import path
from django.contrib import admin
from . import views

app_name = 'main'

urlpatterns = [
    path('', views.ListPost.as_view()),
    path('<int:pk>/',views.DetailPost.as_view())
    
]