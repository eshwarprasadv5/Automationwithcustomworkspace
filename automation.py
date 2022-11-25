from openpyxl import Workbook, load_workbook
import sys
wb = load_workbook('data/testsuite/Testsuite.xlsx')
sh=wb['TestSuite'] 
test_folders_loc=[]
for sh_rowB in sh['B']:
        test_folders_loc.append(sh_rowB.value)
test_suite_val1=[]
for sh_rowA in sh['A']:
        test_suite_val1.append(sh_rowA.value)
input_file_name=list(map(str, sys.argv[1].split(',')))
print("Input file names are : ", input_file_name)
input_product_name=sys.argv[2]
print("Input product name is : ", input_product_name)
for file in input_file_name:
   if(file in test_folders_loc):
    t1=test_folders_loc.index(file)+1
    sh['A'+str(t1)].value=1
    wb.save('data/testsuite/Testsuite.xlsx')
    wb2=load_workbook(file)
    sh2=wb2['Testcases']
    product_names=[]
    for sh2_rowC in sh2['C']:
        product_names.append(sh2_rowC.value)  
    for i in range(len(product_names)):
        if(product_names[i]==input_product_name):
         sh2['A'+str(i+1)].value=1
         wb2.save(file)
   else:
    print('Input file not exists in list')
