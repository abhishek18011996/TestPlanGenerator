from testGeneratorPackage import testgenerate

input_dictionary={'Device Type':None,'Series':None,'Dev type':None,'Team':None}
input_dictionary['Device Type']='Tablet'
input_dictionary['Dev type']='Dev'

testgenerate.generate_test_plan(input_dictionary,'/resources/Transactions.xlsx','TestItems','TeamData')
