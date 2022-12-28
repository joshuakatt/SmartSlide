import collections
import collections.abc
from pptx import Presentation
import glob
import textrazor
import logging
import threading
import time


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


def thread_function(name):
    logging.info("Thread %s: starting", name)

    time.sleep(2)
    logging.info("Thread %s: finishing", name)


if __name__ == "__main__":
    format = "%(asctime)s: %(message)s"
    logging.basicConfig(format=format, level=logging.INFO,
                        datefmt="%H:%M:%S")

    threads = list()
    for index in range(3):
        logging.info("Main    : create and start thread %d.", index)
        x = threading.Thread(target=thread_function, args=(index,))
        threads.append(x)
        x.start()

    for index, thread in enumerate(threads):
        logging.info("Main    : before joining thread %d.", index)
        thread.join()
        logging.info("Main    : thread %d done", index)
