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
                currentRun = para.runs[count+1]
                count+=1
            if(len(text2add) > 0):
                fullText.append(text2add)

    #print(fullText)
    return fullText


def card2numb(img):
    baseurl = "https://db.ygoprodeck.com/api/cardinfo.php?name="
    first_response = get(baseurl + img, timeout=5)
    response_list = first_response.json()
    #print(response_list)
    if("error" in response_list):
        #print("Error: Bad Token: " + img)
        return ""
    else:
        return response_list[0]['id']


if __name__ == '__main__':
    scriptname = input("Script name:")
    scriptwithExt = scriptname + ".docx"
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
            if len(card2numb(img)) > 1:
                print("Downloading " + img + "...")
                url = "https://storage.googleapis.com/ygoprodeck.com/pics/" + card2numb(img) + ".jpg"
                response = get(url, stream=True)
                response.raw.decode_content = True

                with open(fnameImg + fscript + img.replace(":", "") + ".jpg", 'wb') as f:
                    copyfileobj(response.raw, f)

    print("All Done! Feel free to close this window")
    input("")
