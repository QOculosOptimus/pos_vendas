from django.db import models
from django.utils import timezone

class BlingAuth(models.Model):
    """
    Single row table storing oauth tokens.
    """
    # add primary key that is an integer
    number = models.IntegerField(primary_key=True, default=1)
    access_token = models.TextField()
    refresh_token = models.TextField()
    expires_at = models.DateTimeField()
    bling_client_id = models.TextField()
    bling_client_secret = models.TextField()
    bling_redirect_uri = models.TextField()
    bling_oauth_authorize = models.TextField()
    bling_token_endpoint = models.TextField()

    @property
    def is_expired(self):
        return timezone.now() >= self.expires_at
