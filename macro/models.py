from django.db import models

# Create your models here.

class macro(models.Model):
	title = models.CharField(max_length=120)
	content = models.TextField()
	publishing_date = models.DateTimeField()
