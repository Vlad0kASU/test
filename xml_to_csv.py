import os
from progress.bar import IncrementalBar
from time import time
import xml.dom.minidom as md
import xml.etree.ElementTree as ET
import sys
import csv

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

    path = os.getcwd()
    xml=[]
    for dirs, folder, files in os.walk(path):
        for file in files:
            if ".xml" in f"{os.path.join(dirs, file)}".lower():
                # if "venv" not in f"{os.path.join(dirs, file)}".lower() and ".idea" not in f"{os.path.join(dirs, file)}".lower():
                xml.append(os.path.join(dirs, file))
    directory = os.getcwd()
    files = os.listdir(directory)
    bar = IncrementalBar('Прогресс', max=len(xml))

    for file in xml:
        try:
            bar.next()
            backslash="\\"
            print(f' {file[file.rfind(backslash)+1:]}')
            xmldoc = md.parse(file)
            tree = ET.parse(file)
            root = tree.getroot()
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

            Names = []
            # NamesValue = []
            dicts = []
            for i in range(len(Items)):
                p=0
                k=0
                number=0
                if k==0:
                    for item in Items[i]:
                        if len(item.attributes.values())>k:
                            k=len(item.attributes.values())
                            number=item
                if p==0:
                    for num in number.attributes.values():
                        Names.append(num.name)
                        # NamesValue.append([])

                file_name = file.lower()
                file_name = file_name[:-3]
                file_name=file_name+"csv"
                with open(file_name, mode="a", encoding='utf-8') as w_file:
                    file_writer = csv.writer(w_file, delimiter="|", lineterminator="\r\n")
                    file_writer.writerow(Names)

                for item in Items[i]:
                    j=0
                    test_list = []

                    for a in item.attributes.values():

                        if a.name!=Names[j]:
                            for m in range(j, len(Names)):
                                if a.name==Names[m]:
                                    j=m
                                    break
                                else:
                                    # NamesValue[m].append("")
                                    test_list.append("")
                        # NamesValue[j].append(a.value)
                        test_list.append(a.value)
                        j += 1
                    with open(file_name, mode="a", newline='', encoding='utf-8') as w_file:
                        file_writer = csv.writer(w_file, delimiter="|", lineterminator="\r\n")
                        file_writer.writerow(test_list)
        except:
            print("Не удалось обработать файл", file)




if __name__ == '__main__':
    sys.excepthook = excepthook
    # starttime=time()
    main()
    # endtime=time()
    # print(endtime-starttime)