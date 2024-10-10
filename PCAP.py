'''
customers = [{'name':'Alice','age': 25}, {'name':'Bob','age': 35}, {'name':'Charlie','age': 28}]
products = [{'name': 'Laptop', 'units_sold': 1200}, {'name': 'Mouse', 'units_sold': 800}, {'name': 'Keyboard', 'units_sold': 1500}]
top=sorted(products,key=lambda p:p['units_sold'],reverse=True)
print(top[:1])


class A:
    x=""
    def __init__(self,s):
        A.x+=s

obj1=A("Python")
obj2=A("Viper")

print(obj1.x)
print(obj2.x)


def bouns_calculation(salary, rating, absence):
    if rating >= 4.5 and absence < 5:
        return salary * 0.2
    elif (rating >= 3.5 and rating < 4.5) and absence < 10:
        return salary * 0.1
    else:
        return 0

def performer(rating):
    if rating >= 4.5 :
        return "High Performer"
    elif (rating >= 3.5 and rating < 4.5) :
        return "Average Performer"
    else:
        return "Low Performer"

employees = [
    {
        "Employee ID": 101,
        "Name": "Alice Johnson",
        "Department": "HR",
        "Monthly Salary": 5000,
        "Performance Rating": 4.8,
        "Years of Experience": 5,
        "Absences": 3,
        "Overtime Hours": 50
    },
    {
        "Employee ID": 102,
        "Name": "Bob Smith",
        "Department": "Engineering",
        "Monthly Salary": 7000,
        "Performance Rating": 4.2,
        "Years of Experience": 8,
        "Absences": 6,
        "Overtime Hours": 80
    },
    {
        "Employee ID": 103,
        "Name": "Carol Williams",
        "Department": "Sales",
        "Monthly Salary": 6000,
        "Performance Rating": 3.9,
        "Years of Experience": 4,
        "Absences": 2,
        "Overtime Hours": 30
    },
    {
        "Employee ID": 104,
        "Name": "David Brown",
        "Department": "Engineering",
        "Monthly Salary": 6500,
        "Performance Rating": 4.6,
        "Years of Experience": 7,
        "Absences": 1,
        "Overtime Hours": 70
    },
    {
        "Employee ID": 105,
        "Name": "Eve Davis",
        "Department": "HR",
        "Monthly Salary": 5500,
        "Performance Rating": 3.4,
        "Years of Experience": 6,
        "Absences": 12,
        "Overtime Hours": 20
    },
    {
        "Employee ID": 106,
        "Name": "Frank Thomas",
        "Department": "Sales",
        "Monthly Salary": 5800,
        "Performance Rating": 4.9,
        "Years of Experience": 3,
        "Absences": 2,
        "Overtime Hours": 90
    },
    {
        "Employee ID": 107,
        "Name": "Grace Lee",
        "Department": "Marketing",
        "Monthly Salary": 5200,
        "Performance Rating": 3.7,
        "Years of Experience": 2,
        "Absences": 9,
        "Overtime Hours": 40
    }
]

bouns_calculated = list(map(lambda e: bouns_calculation(e["Monthly Salary"], e["Performance Rating"], e["Absences"]), employees))
performerFind = list(map(lambda e: performer(e["Performance Rating"]), employees))
for i in range(len(employees)):
    employees[i]["Bonus"] = bouns_calculated[i]
print(employees)

for i in range(len(employees)):
    employees[i]["Performer"] = performerFind[i]
print(employees)

top_performer = max(employees, key=lambda e: e["Performance Rating"])
print(top_performer)



HighPerformerGroup=list(filter(lambda e:e["Performer"] == "High Performer",employees))
AveragePerformerGroup=list(filter(lambda e:e["Performer"] == "Average Performer",employees))
LowPerformerGroup=list(filter(lambda e:e["Performer"] == "Low Performer",employees))

print(HighPerformerGroup)
print(AveragePerformerGroup)
print(LowPerformerGroup)



print(lambda: 10)


print(str(None)!="None")
print(len('""')==2)

print("Python"+"\n"*len("\n")+"Monty")

print("\\")


print('a' and 4)
print(False and True)
print([8] and 0 and None)
print("one two,three,four:five".split())


print("monty PYTHON".capitalize())
print("monty PYTHON".title())
print("monty PYTHON".lower())
print("monty PYTHON".upper())


l="1.325 0.2 5".split()
s=l[-1]+l[0]
print(s)



print(1 or 'a')
print(0 or [])
print('' or 5 or True)



#inheritance code
class Employee:
    def __init__(self, name, id, dept):
        self.name = name
        self.emp_id = id
        self.department = dept

    def disp(self):
        print("Name:", self.name)
        print("Emp Id:", self.emp_id)
        print("Department:", self.department)

class Engineer(Employee):
    def __init__(self, n, id, d, s, expe):
        super().__init__(n, id, d)
        self.specialization = s
        self.experience = expe

    def display(self):
        super().disp()
        print("Specilization:", self.specialization)
        print("Experience:", self.experience)


Eng1 = Engineer("sunil", 123, "sales", "Finance", 3)
Eng1.display()



class A:
    pass

class B:
    pass

class C(A):
    pass

class D(C):
    pass

print(issubclass(D,C),end='') #Child
print(issubclass(D,A),end='') #Grand Child
print(issubclass(A,A),end='') #Self
print(issubclass(D,B),end='') #No releation
print(issubclass(D,(B,A))) # #Atleast one should be the sub class

print(C.__base__)

import re

def clean_text(text):
    text = re.sub(r'<.*?>', "", text)
    text = re.sub(r'[^a-zA-Z0-9\s]', '', text)
    return text


html_text = "<p> this is  <b> product </b> booming. </p>"
cleaned_text = clean_text(html_text)
print(cleaned_text.title())

import logging

# creAte a logger
logger = logging.getLogger("my_logger")
logger.setLevel(logging.DEBUG)

# create file handler
file_handler = logging.FileHandler("app1.log")

file_handler.setLevel(logging.INFO)

logger.addHandler(file_handler)

# define log msg format
formatter = logging.Formatter('%(asctime)s-%(name)s-%(levelname)s-%(message)s')
file_handler.setFormatter(formatter)

logger.debug("this is a debug msg")
logger.error("user logged in")
logger.warning("file was blocked")
logger.info("Log Check")




# 6. data masking
import regex
def mask_sensitive_info(text):
    pattern=r'\d{4}-\d{3}-\d{3}'
    maskedDate=regex.sub(pattern,"xxx",text)
    return  maskedDate

masked_text = mask_sensitive_info(text_document)
print(masked_text)
'''
import re

# failure codes-- 4xx,5xx
def statusCodeValidation(status_codes):
    #status_codes = [202, 404, 514, 220, 420]
    failure_status_count = 0
    for sc in status_codes:
        if int(sc) in range(400, 600):
            failure_status_count += 1
    return failure_status_count


def ipAddressCheck(ip_address):
    #ip_address = ['192.168.1.1', '192.168.1.2', '192.168.1.3', '192.168.1.1', '192.168.1.2']
    count_ip = {}
    for ip in ip_address:
        if ip in count_ip:
            count_ip[ip] += 1
        else:
            count_ip[ip] = 1
    return count_ip


def regexMatchCheck(text):
    #text=f'192.168.1.1 - - [10/Sep/2023:14:23:45 +0000] "GET /index.html HTTP/1.1" 200 1043'
    #patternForIP=r'(\d{1,3}.\d{3}.\d{1}.\d{1})--\[(\d{2}/[a-zA-Z]{3}]/\d{4}:d{2}:d{2}:d{2} +\d{4}])\]"([A-Z]+)(.+?)HTTP/[\d\.]+"(\d{3})(\d+)'
    regex = r'(\d{1,3}(?:\.\d{1,3}){3}) - - \[(\d{2}/[A-Za-z]{3}/\d{4}:\d{2}:\d{2}:\d{2} \+\d{4})\] "([A-Z]+) (.+?) HTTP/[\d\.]+" (\d{3}) (\d+)'
    match_res=re.match(regex,text)
    ipaddress=match_res.group(1)
    statusCode=match_res.group(5)
    return ipaddress,statusCode

ipList=[]
statusList=[]
for x in open("D:\LeapPython\Logfile.txt",'r'):
    ip,statusCode=regexMatchCheck(x)
    ipList.append(ip)
    statusList.append(statusCode)
print(statusList)
ipAddressDict=ipAddressCheck(ipList)
statusCount=statusCodeValidation(statusList)
print(ipAddressDict)
print(statusCount)