from django.http import HttpResponse
from django.views import View
from turret.client import TurretClient
from django.views.decorators.csrf import csrf_exempt
import json

turret_client = TurretClient('glsmanilasouth2020@gmail.com', 'glsms2020')


@csrf_exempt
def send_welcome_email(request):
    if request.method == "POST":
        # data = json.loads(request.body)
        data = json.loads(request.body)
        turret_client.send_email(data["Name_First"], data["Email"])
        data = {"status": "successfully sent"}
        return HttpResponse(json.dumps(data), content_type="application/json")
