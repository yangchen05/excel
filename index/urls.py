#conding=utf-8
from django.conf.urls import url,include
from django.contrib import admin
from index.views import *
urlpatterns = [
    url(r'^upload/', uploadfile),
    url(r'^downfile/', downfile),
    url(r'^index/', index),

]
