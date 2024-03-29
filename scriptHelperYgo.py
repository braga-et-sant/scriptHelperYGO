import os
from shutil import copyfileobj

from docx import Document
from requests import get



def docScour(filename, color):
    color = color.replace("(", "").replace(")", "")
    colorI = tuple(map(int, color.split(', ')))
    doc = Document(filename)
    fullText = []
    for para in doc.paragraphs:
        for count, run in enumerate(para.runs):
            currentRun = run
            text2add = ""
            skipped = 1
            while (currentRun.font.color.rgb == colorI):
                text2add += currentRun.text
                try:
                    currentRun = para.runs[count+1]
                except:
                    break;
                count+=1
            if(len(text2add) > 0):
                fullText.append(text2add)

    #print(fullText)
    return fullText


def card2numb(img):
    baseurl = "https://db.ygoprodeck.com/api/cardinfo.php?name="
    #print("card2numb:")
    #print(img)
    first_response = get(baseurl + img, timeout=5)
    response_list = first_response.json()
    #print(response_list)
    if("error" in response_list):
        #print("Error: Bad Token: " + img)
        return ""
    else:
        return response_list[0]['id']


def executeImgGet():
    scriptname = input("Script name:")
    scriptwithExt = "script/" + scriptname + ".docx"
    fnameImg = "img/"
    fscript = scriptname + "/"
    color = "(255, 0, 0)"
    data = ""
    imglist = docScour(scriptwithExt, color)

    if not os.path.exists(fnameImg):
        # Create a new directory because it does not exist
        os.mkdir(fnameImg)

    if not os.path.exists(fnameImg + scriptname + "/"):
        # Create a new directory because it does not exist
        os.mkdir(fnameImg + scriptname + "/")

    for img in imglist:
        if os.path.isfile(fnameImg + scriptname + "/" + img + ".jpg"):
            pass
        else:
            #print(img)
            #print(img.strip(" .,()"))
            cardid = card2numb(img.strip(" .,()").replace("’", "'"))
            if len(cardid) > 1:
                print("Downloading " + img + "...")
                url = "https://images.ygoprodeck.com/images/cards/" + cardid + ".jpg"
                response = get(url, stream=True)
                response.raw.decode_content = True

                with open(fnameImg + fscript + img.replace(":", "").replace("?", "") + ".jpg", 'wb') as f:
                    copyfileobj(response.raw, f)

    print("All Done! Feel free to close this window")
    input("")

def executeRedGet():
    scriptname = input("Script name:")
    scriptwithExt = "script/" + scriptname + ".docx"
    color = "(255, 0, 0)"
    imglist = docScour(scriptwithExt, color)
    print(imglist)

if __name__ == '__main__':
    executeImgGet()
    #executeRedGet()
