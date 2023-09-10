from ppadb.client import Client

adb = Client(host="127.0.0.1", port=5037)
device = adb.devices()[0]

result = device.screencap()
with open("screen.png", "wb") as fp:
    fp.write(result)