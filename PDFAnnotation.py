# -*- coding: utf-8 -*-
"""
Created on Tue Aug 17 11:27:37 2021

@author: VR787FC
"""

import fitz,re

doc = fitz.open(r"C:\Users\VR787FC\Desktop\SSG.pdf")

NumberPageDict={}
for pgnum,page in enumerate(doc):
    ### SEARCH
    NumberList = re.findall("\d{1,}[\d\,\.]*(?<![\.\,])", page.get_text("text"))
    NumberList= [item for item in set(NumberList) if len(item)>1]  #need to remove . at the end
    for numbers in NumberList:
        if NumberPageDict.get(numbers)==None:
            NumberPageDict[numbers] = [pgnum+1]
        else:
            NumberPageDict[numbers].append(pgnum+1)
        
for pgnum,page in enumerate(doc):
    ### SEARCH
    #"\d{2,}(?:\,\d{3})?(?:\,\d{3})?(?:\.\d{1,2})?"
    NumberList = re.findall("\d{1,}[\d\,\.]*(?<![\.\,])", page.get_text("text"))
    ExcludeList = ["2019","2020","31","19","20"]
    NumberList= [item for item in set(NumberList) if len(item)>1 and item not in ExcludeList]  #need to remove . at the end
    print(NumberList)
    
    PageLocationDict = {}
    for text in NumberList:
        if not any(text in s and len(s)>len(text) for s in NumberList):
            text_instances = page.searchFor(text)
            # LocationList = []
            for location in text_instances:
                # LocationList.append(location)
                PageLocationDict[location] = ','.join(["/{}".format(n) for n in NumberPageDict[text] if n != pgnum+1])
    
            ### HIGHLIGHT
    for inst in PageLocationDict:
        if len(PageLocationDict[inst])>1:
            page.add_freetext_annot((inst[2]+2,inst[1]+1,inst[2]+62,inst[3]+1), PageLocationDict[inst], fontsize=7,text_color= (1, 0, 0) )#blue(0, 0, 1)
            highlight = page.addHighlightAnnot(inst)#(x1,y1,x2,y2)
            highlight.setColors({"stroke":(0, 0.8, 1)})
            highlight.update()
                # if len(NumberPageDict[text])>1:
                #     page.add_freetext_annot((inst[2]+2,inst[1]+1,inst[2]+62,inst[3]+1), ','.join(["/{}".format(n) for n in NumberPageDict[text] if n != pgnum+1]), fontsize=7,text_color= (1, 0, 0) )#blue(0, 0, 1)
                #     highlight = page.addHighlightAnnot(inst)#(x1,y1,x2,y2)
                #     highlight.setColors({"stroke":(0, 0.8, 1)})
                #     highlight.update()
        else:
            highlight = page.addHighlightAnnot(inst)#(x1,y1,x2,y2)
            highlight.setColors({"stroke":(1, 0.8, 0)})
            highlight.update()                    
    # if pgnum+1==1:break

### OUTPUT
doc.save("output.pdf", garbage=4, deflate=True, clean=True)
