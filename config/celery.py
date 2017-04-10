from __future__ import absolute_import, unicode_literals
from celery import Celery, task

app = Celery(broker='redis://localhost:6379/0', include=['utils.upload_ops'])

if __name__ == '__main__':
    app.start()