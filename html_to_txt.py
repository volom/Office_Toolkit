# -*- coding: utf-8 -*-
"""
Created on Mon Feb  8 14:44:44 2021
function for convertin html to text
@author: volom
"""
import urllib.request
from bs4 import BeautifulSoup
def html_to_txt(url):
    html = urllib.request.urlopen(url).read()
    soup = BeautifulSoup(html)
    
    # kill all script and style elements
    for script in soup(["script", "style"]):
        script.extract()    # rip it out
    
    # get text
    text = soup.get_text()
    
    # break into lines and remove leading and trailing space on each
    lines = (line.strip() for line in text.splitlines())
    # break multi-headlines into a line each
    chunks = (phrase.strip() for line in lines for phrase in line.split("  "))
    # drop blank lines
    text = '\n'.join(chunk for chunk in chunks if chunk)
    return text







