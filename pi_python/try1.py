
'''
1)Change throttle time to 5 ms
2)Create an excel sheet and write the rtd function of all the entered scripts (Functions)
3)define function for reading an rtd quote..
4)Create n excel sheets for n different quotes entered in parameters..


                    Flow 
1           Read n quotes vector with 10 values = 10*n fields 

2           Function should independently look for changes by comparing strings instantaneously
        Function must be inside an infinite loop
            Individual functions must contain an incrementer for writing in order inside an excel file


1.0)create excel file                           #Write the RTD data.

2.0)create excel sheets for all the scripts     #Storing data 

main function(infinite loop)

    1.0)sub function_read
        read each val from main excel file
        return String1 

    1.1)sub function_read
        read each val from main excel file
        return String2

    2)sub compare_function(s) 
        Compare String2 & String 1
        return 1 if True//Duplicate no changes
        return 0 //Change detected

    3)sub function_write(Quote,Sheet)
        writes into another excel respective excel file

'''
Quote1 = '"BANKNIFTY"'
Quote2 = '"BANKNIFTYs"'

def Read1(stra):
    x = str(1)
    y = str(3)
    z = str(6)
    String1 = x+y+z
    return String1

def Read2(stra):
    #reading
    x = str(1)
    y = str(3)
    z = str(6)
    String2 = x+y+z
    return String2

def Compare(String1,String2):
    if String1==String2:
        return 1
    else:
        return 0

String1 = Read1("x")
String2 =Read2("y")
x = Compare(String1,String2)
print(x)






