import urllib.request
url = "https://github.com/Mr-cool08/hembrev/raw/master/hembrev.exe"
print ("download start!")
filename, headers = urllib.request.urlretrieve(url, filename="/program/hembrev.exe")
print ("download complete!")
print ("download file location: ", filename)
print ("download headers: ", headers)