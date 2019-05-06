import wmi
import time 
c = wmi.WMI()
while True:
  for service in c.Win32_Service(Name="IFaceAttReader"):
    result, = service.StartService()
    if result == 0:
      print("Service", service.Name, "started")
    elif result == 10:
      print("service is already running.")
    else:
      print("Some problem", result)
    break
  else:
    print("Service not found")
  time.sleep(300)