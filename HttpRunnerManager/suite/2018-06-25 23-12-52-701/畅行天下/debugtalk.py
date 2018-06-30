import hashlib
import hmac
import json
import os
import random
import string
import time
import base64



def setup_hook_add_kwargs(request):
   print("HHHHHHHHHHHHHHHHHHHHHHHHHHHHHH")
   if request["method"] == "POST":
        print("AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA")
        print(request["data"]["id"])
        print(request["data"]["name"])
        # request["data"] = json.dumps(request["data"])
        bytesString = json.dumps(request["data"]).encode(encoding="utf-8")
        print(bytesString)

        #base64 编码
        request["data"] = base64.b64encode(bytesString)
        print(request["data"])

def setup_hook_remove_kwargs(request):
    print("====================================================================")

def teardown_hook_sleep_N_secs(response, n_secs):
    print("oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo")
    """ sleep n seconds after request
    """
    # if response.status_code == 200:
    #     time.sleep(0.1)
    # else:
    #     time.sleep(n_secs)
