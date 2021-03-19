import urllib.request
import urllib3
import os
import time

p_m3u8 = input("Enter m3u8 path in txt format:")
p_http = input("Enter url of ts files (only the part before the last slash):")
p_output = input("Enter path of output folder")
p_output_decrypted = input("Enter path of descrypted output folder")
p_output_combined = input("Enter path of combined output folder")
IV = input("Enter IV (search in the first few lines of m3u8 file opened in notepad, do not include 0x):")
K = input("Enter Key (search in the first few lines of m3u8 file opened in notepad for url of the key file,"
          "and convert the file into hex with the help of online tools):")


def download_tsfiles(p_m3u8, p_http, p_output):
    opener = urllib.request.build_opener()
    opener.addheaders = [('User-Agent',
                          'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/'
                          '537.36 (KHTML, like Gecko) Chrome/36.0.1941.0 Safari/537.36')]
    urllib.request.install_opener(opener)
    existing_files = os.listdir(p_output)
    with open(p_m3u8, "r") as m3u8:
        for line in m3u8.readlines():
            line = line.strip("\n")
            if line.endswith("ts") and line not in existing_files:
                url_ts = p_http + "/" + line
                print("downloading " + str(url_ts))
                while True:
                    try:
                        print("trying")
                        urllib.request.urlretrieve(url_ts, filename=p_output+"\\"+line)
                        break
                    except urllib3.exceptions.HTTPError as err:
                        print(err)
                        time.sleep(300)


def decryption(p_output, p_output_decrypted):
    for file in os.listdir(p_output):
        p_output_decrypted_file = os.path.join(p_output_decrypted, file)
        open(p_output_decrypted_file, "w")
        cmd1 = "cd C:\\Program Files\\OpenSSL-Win64\\bin&&"
        cmd2 = "openssl aes-128-cbc -d -in " + p_output + "\\" + file + " -out " + p_output_decrypted_file + \
               " -nosalt -iv " + IV + " -K " + K
        os.system(cmd1+cmd2)


def combine_files(p_output_decrypted, p_output_combined):
    with open(p_output_combined, "wb") as f:
        with open(p_m3u8, "r") as m3u8:
            for line in m3u8.readlines():
                line = line.strip("\n")
                if line.endswith("ts"):
                    print(line)
                    ts_video_path = os.path.join(p_output_decrypted, line)
                    f.write(open(ts_video_path, 'rb').read())


download_tsfiles(p_m3u8, p_http, p_output)
decryption(p_output, p_output_decrypted)
combine_files(p_output_decrypted, p_output_combined)
