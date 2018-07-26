import inspect

#import the module
from xlsx_converter_to_django_class import convertXlsxToDjangoClass

#define the class we want to use
class DjangoClass:
    alluminium = ""
    clientId = ""
    clientContact = ""
    clientReference = ""
    dateReceived = ""
    reportNo = ""
    nitrate = ""
    nitrite = ""
    sampleRef = ""
    sampleReportNumber = ""
    sampleID = ""
    totalColiforms = ""

#define the path for the file to read
filePathExample = "./Sample-data-3.xlsx"
 
#get the data with the library
parsedData = convertXlsxToDjangoClass(DjangoClass, filePathExample, headerLimit=30)

#show the results
for elem in parsedData:
        print("element:")
        for header in [b[0] for b in [a for a in inspect.getmembers(DjangoClass, lambda a:not(inspect.isroutine(a))) if not(a[0].startswith('__') and a[0].endswith('__'))]]:
            print("\t{0}: {1}".format(header, getattr(elem, header)))