import collections
import collections.abc
from pptx import Presentation
import glob
import textrazor


pptData = []


razor_API_KEY = "9224bb6f01f35cb07f2ec8b1c7ca780f397685e30b95d517a3eaeaee"
textrazor.api_key = razor_API_KEY
client = textrazor.TextRazor(extractors=["entities", "topics"])

for eachfile in glob.glob("text_test1.pptx"):
    prs = Presentation(eachfile)
    print(eachfile)
    print("----------------------")
    pptData = [""]*len(prs.slides)
    for x in range(len(prs.slides)):
        for shape in prs.slides[x].shapes:
            if hasattr(shape, "text"):
                pptData[x] = shape.text
                print(shape.text)
    print(pptData)

print("----------------------------------------------------------------------------------")
print("----------------------------------------------------------------------------------")
print("----------------------------------------------------------------------------------")
print("----------------------------------------------------------------------------------")
for x in range(len(pptData)):
    try:
        response = client.analyze(pptData[x])
        json_content = response.json
        new_response = textrazor.TextRazorResponse(json_content)
        entities = list(response.entities())
        topic = response.topics
        entities.sort(key=lambda x: x.relevance_score, reverse=True)
        seen = set()
        for entity in entities:
            if entity.id not in seen:
                print(x, entity.wikipedia_link)
    except:
        pass
