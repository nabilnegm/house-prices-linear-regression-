import numpy as np
from sklearn.linear_model import LinearRegression
import xlrd
import xlwt
import csv


#load the train sheet values into a normal array
trainbook = xlrd.open_workbook('train.xlsx')
trainsheet = trainbook.sheet_by_name('Worksheet')
train_normal = []
for i in range (1461):
    train_normal.append([])
    for j in range (81):
        train_normal[i].append(trainsheet.cell(i, j).value)

#put the values in numpy array
train_numpy = np.array(train_normal)

#PREPROCESSING OF DATA
#change all strings to numbers
for j in range (len(train_numpy[1])):

    #check if this coloumn is strings
    if len(train_numpy[2][j]) == 1 :

        if not train_numpy[2][j][0].isdigit() :

            temp = []
            #append all strings in temp except NA
            for i in range(len(train_numpy)):


                if train_numpy[i][j] == 'NA':
                    # print('el code yeee')

                    train_numpy[i][j] = 0.001
                    continue
                if train_numpy[i][j] in temp or i == 0:
                    continue

                else:
                    temp.append(train_numpy[i][j])

            #change strings to numbers
            for i in range(len(train_numpy)):
                if i == 0:
                    continue
                for k in range(len(temp)):
                    if temp[k] == train_numpy[i][j]:
                        train_numpy[i][j] = k


            else:
                for i in range(len(train_numpy)):
                    if i == 0:
                        continue
                    if train_numpy[i][j] == 'NA':
                        train_numpy[i][j] = 0.001
                        continue
                    else:
                        train_numpy[i][j] = int(float(train_numpy[i][j]))

    else :
        if not train_numpy[2][j][1].isdigit() :

            temp = []
            #append all strings in temp except NA and change NA to a random small variable 0.001
            for i in range(len(train_numpy)):

                if train_numpy[i][j] == 'NA':
                    train_numpy[i][j] = 0.001
                    continue

                if train_numpy[i][j] in temp or i == 0:
                    continue

                else:
                    temp.append(train_numpy[i][j])

            #change strings to numbers
            for i in range(len(train_numpy)):
                if i == 0:
                    continue
                for k in range(len(temp)):
                    if temp[k] == train_numpy[i][j]:
                        train_numpy[i][j] = k

        else:
            for i in range(len(train_numpy)):
                if i == 0:
                    continue
                if train_numpy[i][j] == 'NA':
                    train_numpy[i][j] = 0.001
                    continue
                else:
                    train_numpy[i][j] = int(float(train_numpy[i][j]))





#remove the headers and the ids
train_data = np.array(train_numpy[1: , 1:]).astype(np.float)


# change the NA  which is now 0.001 to the mean in each coloumn
for j in range (len(train_data[1])):

    avg = np.mean(train_data[:,j], axis = 0)
    a = np.where(train_data[:,j] == 0.001)
    for i in range (len(a[0])):
        train_data[a[0],j] = avg


#train_x and train_y are the train_data
train_x = np.array(train_data[: , : -1])
train_y = np.array(train_data[: , -1])


#TRAIN THE MODEL ON train_data

reg = LinearRegression().fit(train_x, train_y)
# print(reg.score(train_x, train_y))




#load the test sheet values into a normal array
testbook = xlrd.open_workbook('test.xlsx')
testsheet = testbook.sheet_by_name('Worksheet')


test_normal = []
for i in range (1460):
    test_normal.append([])
    for j in range (80):
        test_normal[i].append(testsheet.cell(i, j).value)


#put the values in numpy array
test_numpy = np.array(test_normal)


#PREPROCESSING OF DATA
#change all strings to numbers
for j in range (len(test_numpy[1])):

    #check if this coloumn is strings
    if len(test_numpy[2][j]) == 1 :

        if not test_numpy[2][j][0].isdigit() :

            temp = []
            #append all strings in temp except NA
            for i in range(len(test_numpy)):


                if test_numpy[i][j] == 'NA':
                    # print('el code yeee')

                    test_numpy[i][j] = 0.001
                    continue
                if test_numpy[i][j] in temp or i == 0:
                    continue

                else:
                    temp.append(test_numpy[i][j])

            #change strings to numbers
            for i in range(len(test_numpy)):
                if i == 0:
                    continue
                for k in range(len(temp)):
                    if temp[k] == test_numpy[i][j]:
                        test_numpy[i][j] = k


            else:
                for i in range(len(test_numpy)):
                    if i == 0:
                        continue
                    if test_numpy[i][j] == 'NA':
                        test_numpy[i][j] = 0.001
                        continue
                    else:
                        test_numpy[i][j] = int(float(test_numpy[i][j]))

    else :
        if not test_numpy[2][j][1].isdigit() :

            temp = []
            #append all strings in temp except NA and change NA to a random small variable 0.001
            for i in range(len(test_numpy)):

                if test_numpy[i][j] == 'NA':
                    test_numpy[i][j] = 0.001
                    continue

                if test_numpy[i][j] in temp or i == 0:
                    continue

                else:
                    temp.append(test_numpy[i][j])

            #change strings to numbers
            for i in range(len(test_numpy)):
                if i == 0:
                    continue
                for k in range(len(temp)):
                    if temp[k] == test_numpy[i][j]:
                        test_numpy[i][j] = k

        else:
            for i in range(len(test_numpy)):
                if i == 0:
                    continue
                if test_numpy[i][j] == 'NA':
                    test_numpy[i][j] = 0.001
                    continue
                else:
                    test_numpy[i][j] = int(float(test_numpy[i][j]))




#remove the headers and the ids
test_data = np.array(test_numpy[1: , 1:]).astype(np.float)


# change the NA  which is now 0.001 to the mean in each coloumn
for j in range (len(test_data[1])):

    avg = np.mean(test_data[:,j], axis = 0)
    a = np.where(test_data[:,j] == 0.001)
    for i in range (len(a[0])):
        test_data[a[0],j] = avg




output = np.array(test_numpy[0:,[0,1]])
output[0][1] = 'SalePrice'

prediction = reg.predict(test_data)


#put the prediction in the output array
for i in range (len(output)):
    if i == 0 :
        continue
    else:
        output[i][1] = prediction[i-1]




#save the output array into an excel sheet

salesprice = open('saleprice.csv', 'w')
wr = csv.writer(salesprice, quoting=csv.QUOTE_ALL)

for i in range (len(output)):
    if i == 0:
        wr.writerow(output[i])
    else:
        wr.writerow([int(float(output[i][0])),int(float(output[i][1]))])


salesprice.close()
