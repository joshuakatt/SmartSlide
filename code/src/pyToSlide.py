import collections
import collections.abc
from pptx import Presentation
import glob
import textrazor
import logging
import threading
import time
from pptx import Presentation

prs = Presentation('text_test.pptx')

f = open('text_test.pptx')
