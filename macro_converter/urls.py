"""macro_converter URL Configuration

The `urlpatterns` list routes URLs to views. For more information please see:
    https://docs.djangoproject.com/en/2.0/topics/http/urls/
Examples:
Function views
    1. Add an import:  from my_app import views
    2. Add a URL to urlpatterns:  path('', views.home, name='home')
Class-based views
    1. Add an import:  from other_app.views import Home
    2. Add a URL to urlpatterns:  path('', Home.as_view(), name='home')
Including another URLconf
    1. Import the include() function: from django.urls import include, path
    2. Add a URL to urlpatterns:  path('blog/', include('blog.urls'))
"""

from django.conf.urls import url
from django.contrib import admin
from django.urls import path

from macro.views import *

urlpatterns = [
    path('admin/', admin.site.urls),
	url(r'^$', home_view),
    url(r'^about/$', about),
    url(r'^benefits/$', benefits),
    url(r'^progress/$', progress),
    url(r'^used-technologies/$', used_technologies),
    url(r'^license/$', license),
    url(r'^examples/$', examples),
    url(r'^contact/$', contact),
    url(r'^contribute/$', contribute),
	url(r'^test/$', test),
]
