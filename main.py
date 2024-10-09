import os, sys, platform, time, csv, openpyxl
from colorama import Fore, init
from vobject import vCard
init(autoreset=True)
RD = Fore.RED
BL = Fore.BLUE
GR = Fore.GREEN
YL = Fore.YELLOW
WH = Fore.WHITE
CY = Fore.CYAN
RS = Fore.RESET

def Banner():
  return f"""
{BL}        /$$$$$ /$$   /$$         /$$   /$$ /$$
       |__  $$| $$  | $$        | $$$ | $$|__/
          | $$| $$  | $$        | $$$$| $$ /$$ /$$   /$$
          | $$| $$  | $$ /$$$$$$| $$ $$ $$| $$|  $$ /$$/
     /$$  | $$| $$  | $$|______/| $$  $$$$| $$ \  $$$$/
    | $$  | $$| $$  | $$        | $$\  $$$| $$  >$$  $$
    |  $$$$$$/|  $$$$$$/        | $$ \  $$| $$ /$$/\  $$
     \______/  \______/         |__/  \__/|__/|__/  \__/{RS}

{GR}          /$$$$$$  /$$    /$$
         /$$__  $$| $$   | $$
        | $$  \__/| $$   | $$
        | $$      |  $$ / $$/
        | $$       \  $$ $$/
        | $$    $$  \  $$$/  {RD}Author : ./JU-Nix
{GR}        |  $$$$$$/   \  $/
         \______/     \_/
"""



class CV:
  def __init__(self, path, name):
    self.path = path
    self.name = name
  def txttovcf(self):
    with open(self.path, 'r') as f:
      lines = f.readlines()

    arr_f = []
    name, number = None, None
    for line in lines:
        if line.startswith("name:"):
            name = line.split(":")[1].strip()
        elif line.startswith("number:"):
            number = line.split(":")[1].strip()

        if name and number:
            vcard = vCard()
            vcard.add('fn').value = name

            tel = vcard.add('tel')
            tel.value = number
            tel.type_param = 'cell'
            arr_f.append(vcard)

    with open(f'/sdcard/{self.name}.vcf', 'w') as vcf:
        for ctc in arr_f:
            vcf.write(ctc.serialize())

  def csvtovcf(self):
    with open(self.path, 'r') as f:
      csv_reader = csv.DictReader(f)


      with open(f'/sdcard/{self.name}.vcf', 'w') as vcf:
        for row in csv_reader:
          vcard = vCard()

          if row.get('name'):
            vcard.add('fn').value = row['name']

          if row.get('number'):
            tel = vcard.add('tel')
            tel.value = row['number']
            tel.type_param = 'cell'

          vcf.write(vcard.serialize())
          vcf.write('\n')

  def xlsxtovcf(self):
    workbook = openpyxl.load_workbook(self.f_import)
    sheet = workbook.active
    with open(f'/sdcard/{self.name}.vcf', 'w') as vcf:
      for row in sheet.iter_rows(min_row=1, values_only=True):
        name, number = row

        name = str(name) if name else ''
        number = str(number) if number else ''

        vcard = vCard()

        if name and number:
          vcard.add('fn').value = name

          tel = vcard.add('tel')
          tel.value = number
          tel.type_param = 'cell'

        vcf.write(vcard.serialize())
        vcf.write('\n')

def clear_term():
  if platform.system() == "Windows":
    os.system('cls')
  else:
    os.system('clear')

def insert_file(type):
  clear_term()
  print(Banner())
  jum=int(input(f'{YL} Jumlah File (Max 5) '))
  for _ in range(jum):
    if type == "TXT":
      file_path = input(f'{GR} {_} Insert file: ')
      file_name = input(f'{GR} {_} Insert name: ')
      Convert = CV(file_path, file_name)
      Convert.txttovcf()
    if type == "CSV":
      file_path = input(f'{GR} {_} Insert file: ')
      file_name = input(f'{GR} {_} Insert name: ')
      Convert = CV(file_path, file_name)
      Convert.csvtovcf()
    if type == "XLSX":
      file_path = input(f'{GR} {_} Insert file: ')
      file_name = input(f'{GR} {_} Insert name: ')
      Convert = CV(file_path, file_name)
      Convert.xlsxtovcf()

def menusplit():
  while True:
    clear_term()
    print(Banner())
    print(f'{GR} Select Options \n{WH} 1 {CY} VCF \n{WH} 2 {CY} CSV \n{WH} 3 {CY} XLS/X \n')
    select=input(f'{YL} |> ')
    if select == "0":
      break

def menucv():
  while True:
    clear_term()
    print(Banner())
    print(f'{GR} Select Options \n{WH} 1 {CY} TXT To VCF \n{WH} 2 {CY} CSV To VCF \n{WH} 3 {CY} XLS/X To VCF \n')
    select=input(f'{YL} |> ')
    if select == "0":
      break
    elif select == "1":
      insert_file(type="TXT")
    elif select == "2":
      insert_file(type="CSV")
    elif select == "3":
      insert_file(type="XLSX")


def main():
  while True:
    clear_term()
    print(Banner())
    print(f'{WH} Notes:{RD} split is still in development')
    print(f'{GR} Select Options \n{WH} 1 {CY} Cv \n{WH} 2 {CY} Split \n')
    select=input(f'{YL} |> ')

    if select == "0":
      clear_term()
      break
    elif select == "1":
      menucv()
    elif select == "2":
      menusplit()


if __name__ == '__main__':
  main()
