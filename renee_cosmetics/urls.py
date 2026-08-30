"""
URL configuration for renee_cosmetics project.

The `urlpatterns` list routes URLs to views. For more information please see:
    https://docs.djangoproject.com/en/6.0/topics/http/urls/
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

from django.conf import settings
from django.conf.urls.static import static
from django.contrib import admin
from django.http import HttpResponse
from django.templatetags.static import static as static_url
from django.urls import include, path
from django.views.generic.base import RedirectView

urlpatterns = [
    path("admin/", admin.site.urls),
    # Liveness probe — touches nothing (no DB, no auth). Ping it from a FREE
    # external scheduler every ~10 min during work hours to keep the Render free
    # dyno warm (kills cold-start + the ~700ms TiDB TLS handshake on first hit).
    path('healthz', lambda r: HttpResponse('ok', content_type='text/plain')),
    # Browsers probe /favicon.ico at the domain root — point it at our icon
    # so it resolves instead of logging a 404.
    path('favicon.ico', RedirectView.as_view(url=static_url('core/favicon.ico'), permanent=True)),
    path('', include('core.urls')),
    path('offline/', include('offline.urls')),
    path('b2b/', include('online_b2b.urls')),
    path('grn/', include('grn.urls')),
]

if settings.DEBUG:
    urlpatterns += static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)
