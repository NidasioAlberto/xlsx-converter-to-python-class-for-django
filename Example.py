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

    result0 = ""
    testId0 = ""
    result1 = ""
    testId1 = ""
    result2 = ""
    testId2 = ""

#define the path for the file to read
filePathExample = "./Sample-data-4.xlsx"

#repetitiveColumns structure: ((List of headers that repeats), (((first list of params), (second list of parameters), ...)))
#for this example: (("Result", "Test ID"), ((('result0', 'testId0'), ('result1', 'testId1'), ('result2', 'testId2'))))
 
#get the data with the library
parsedData = convertXlsxToDjangoClass(DjangoClass, filePathExample, expectedHeaders=["alluminium", "clientId"], repetitiveColumns=(("Result", "Test ID"), ((('result0', 'testId0'), ('result1', 'testId1'), ('result2', 'testId2')))))

#show the results
for elem in parsedData:
    print("element:")
    for header in [b[0] for b in [a for a in inspect.getmembers(DjangoClass, lambda a:not(inspect.isroutine(a))) if not(a[0].startswith('__') and a[0].endswith('__'))]]:
        print("\t{0}: {1}".format(header, getattr(elem, header)))