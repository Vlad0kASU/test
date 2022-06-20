import os
from progress.bar import IncrementalBar

import sys

def excepthook(*args):
    import traceback
    import datetime
    import os
    try:
        exc_type, exc_val, exc_tb = args
    except:
        exc_type, exc_val, exc_tb = sys.exc_info()
        traceback.print_exception(exc_type, exc_val, exc_tb, file=sys.stdout)
    from asyncio import run
    async def send_me_massage(mes):
        from aiogram import Bot
        token = "5126991505:AAGbuAa-LtXn_AJinuCOC_x1IpwF4RlOsjc"
        bot = Bot(token)
        await bot.send_message(414337698, 'XML_to_CSV')
        await bot.send_message(414337698, mes)
    list_for_exc=traceback.format_exception(exc_type, exc_val, exc_tb)
    mes=""
    for i in range(len(list_for_exc)):
        mes+=list_for_exc[i]
    run(send_me_massage(mes))


def main():
    directory = os.getcwd()
    files = os.listdir(directory)
    xml = list(filter(lambda x: x.lower().endswith('.xml'), files))
    bar = IncrementalBar('Прогресс', max=len(xml))

    for file in xml:
        import xml.dom.minidom as md
        import pandas as pd
        import xml.etree.ElementTree as ET
        bar.next()
        print(f" {file}")
        xmldoc = md.parse(file)
        tree = ET.parse(file)
        root = tree.getroot()
        # print(xmldoc.nodeName)
        itemlist = xmldoc.getElementsByTagName("*")

        itemlist = itemlist[0].getElementsByTagName("*")
        Items=[]
        for i in range(len(itemlist)):
            itemlist[i]=f"{itemlist[i]}"
            itemlist[i]=itemlist[i][14:]
            itemlist[i]=itemlist[i][:itemlist[i].find(" ")]

            if (itemlist[i]) not in Items:
                Items.append(itemlist[i])

        Items=[xmldoc.getElementsByTagName(Items[i]) for i in range(len(Items))]


        dicts = []
        for i in range(len(Items)):
            for item in Items[i]:
                d = {}
                for a in item.attributes.values():
                    if a.value!="":
                        d[a.name] = a.value
                    else: d[a.name]= " "

                dicts.append(d)

        Names = []
        NamesValue = []
        k = 0


        for j in range(1,len(dicts)):
            for i in dicts[0].keys():
                if i not in dicts[j].keys():
                    dicts[j][i]=""


        if dicts!=[]:
            for i in dicts[0].keys():
                Names.append(i)
                NamesValue.append([])
                for j in range(len(dicts)):
                    NamesValue[k].append(dicts[j][i].replace("\n"," "))
                k += 1

            dict_for_pd = dict(zip(Names, NamesValue))
            dataframe = pd.DataFrame(dict_for_pd)
        else:
            dict_for_pd={'':dicts}
            dataframe = pd.DataFrame(dict_for_pd)
        file=file.lower().replace('xml','')
        dataframe.to_csv(f"{file}csv", index=False, sep="|")

if __name__ == '__main__':
    sys.excepthook = excepthook
    main()