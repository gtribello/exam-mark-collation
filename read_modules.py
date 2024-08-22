import pandas as pd
import numpy as np
import glob
import math
import json
import os

def getMaximumMark( mvalue ) :
    if "resit_mark" not in mvalue.keys() : 
       if np.isnan(float(mvalue["mark"])) : return 0 
       else : return float(mvalue["mark"])
    if np.isnan(float(mvalue["mark"])) : maxmark = 0
    else : maxmark = float(mvalue["mark"])
    for m in mvalue["resit_mark"] :
        mm  = float(m)
        if not float(np.isnan(mm)) and mm>maxmark : maxmark = mm
    return maxmark
    
def getMark( row, data, studentno,) :
    if data.isnull()["MarkResult"].at[row.Index] : 
       return { "mark": "U", "EBN": "U" } 
    elif data["MarkResult"].at[row.Index]=="P" or "PAL" in data["MarkResult"].at[row.Index] or "PAS" in data["MarkResult"].at[row.Index] or data.isnull()["ResitResult"].at[row.Index] :
       return {"mark": data["Mark"].at[row.Index], "EBN": data["MarkResult"].at[row.Index] }
    return {"mark": data["Mark"].at[row.Index], "EBN": data["MarkResult"].at[row.Index], "resit_mark": [data["Resit"].at[row.Index]], "resit_EBN": [data["ResitResult"].at[row.Index]] }

def addRecordToDict( row, data, output ) :
   module = data["Subject"].at[row.Index] + str(data["Catalog"].at[row.Index])
   if module not in output.keys() :
      output[module] = getMark( row, data, output["studentno"] )
   else :
      saved_result = output[module]
      output[module] = getMark( row, data, output["studentno"] )
      if "resit_mark" not in output[module].keys() : output[module]["resit_mark"], output[module]["resit_EBN"] = [], []
      output[module]["resit_mark"].append( saved_result["mark"] )
      output[module]["resit_EBN"].append( saved_result["EBN"] )
      if "resit_EBN" in saved_result.keys() :
         for i in range(len(saved_result["resit_mark"])) : 
             output[module]["resit_mark"].append( saved_result["resit_mark"][i] )
             output[module]["resit_EBN"].append( saved_result["resit_EBN"][i] )

def getFinalMark( mvalue ) :
   if mvalue["EBN"]=="P" or "PAL" in mvalue["EBN"] or "PAS" in mvalue["EBN"] : return mvalue["mark"]
   if mvalue["EBN"]=="PH" : return 40
   if "resit_EBN" in mvalue.keys() :
      if mvalue["resit_EBN"][-1]=="PH" : return 40
      if mvalue["resit_EBN"][-1]=="P" : return mvalue["resit_mark"][-1]
      return getMaximumMark( mvalue )
   if "ABSM" in mvalue["EBN"] or "AbsM" in mvalue["EBN"] : return "ABSM"
   if "ABS" in mvalue["EBN"] : return 0
   if mvalue["EBN"]=="U" : return "?"
   if "F" not in mvalue["EBN"] : print("found curious mark", mvalue )
   return mvalue["mark"]

def getFinalModuleMarks( value, handbook ) :
   outdict = {}
   for mkey, mvalue in value.items() :
       if mkey not in handbook.keys() : outdict[mkey] = mvalue 

   for mkey, mvalue in value.items() :
       if mkey not in handbook.keys() : 
          continue  
       else : outdict[mkey] = getFinalMark( mvalue )
   return outdict

def check_predominance( target, value, handbook ) :
   nl1, nl2, nl3 = 0, 0, 0
   for mkey, mvalue in value.items() :
       if mkey not in handbook.keys() : continue 

       if float(getFinalMark( mvalue ))>=target : 
          if handbook[mkey]["level"]==1 : nl1 = nl1 + handbook[mkey]["CATS"]/20
          elif handbook[mkey]["level"]==2 : nl2 = nl2 + handbook[mkey]["CATS"]/20
          else : nl3 = nl3 + handbook[mkey]["CATS"]/20

   if 10*nl1 + 30*nl2 + 60*nl3>=300 : return "yes"
   return "no"

# Check if there is a completed file
completed_students = []
if os.path.exists("completed.xlsx") :
   complete_data = pd.read_excel("completed.xlsx")
   completed_students = complete_data["ID"].tolist()

# Read in the excel 1st sheet of the excel file that contains the student modules
data = pd.read_excel('QSR_EXAM_RESULTS_1379.xlsx')

# Read the dictionary containing the list of modules that we understand
with open("../handbook.json") as h : handbook = json.load(h)

students = {}
for row in data.itertuples() :
   studentno = data['ID'].at[row.Index]
   if studentno in completed_students : continue
   module = data["Subject"].at[row.Index] + str(data["Catalog"].at[row.Index])

   # Check we know what this module is
   if module not in handbook.keys() : raise Exception("Found module " + module + " that is not in handbook")
   # Skip it if it is a zero points module or a placement year
   if handbook[module]["CATS"]==0 or handbook[module]["CATS"]==120 : continue

   if studentno not in students.keys() : 
      localdict = {} 
      localdict["studentno"] = int(studentno)
      localdict["name"] = name = data['First Name'].at[row.Index] + " " + data['Last'].at[row.Index]
      localdict["startyear"] = str(data['Admit Term'].at[row.Index])
      localdict["degprogram"] = data['Plan Descr'].at[row.Index].split()[0]
      addRecordToDict( row, data, localdict )
      students[int(studentno)] = localdict
   else :
      addRecordToDict( row, data, students[studentno] )

# Now read in all the level three modules
tabname = "Supplementary"   # "Main"    # Specify which tab of the grade roster to read in
for module in list(glob.glob('*.xlsx')):
    if module == 'QSR_EXAM_RESULTS_1379.xlsx' or module == "completed.xlsx" : continue
    modulename = module.replace(".xlsx","")
    print("GETTING DATA FOR MODULE", module )
    module_df = pd.read_excel(module,tabname,skiprows=39)
    for row in module_df.itertuples() :
        studentno = module_df["Student No"].at[row.Index]
        if studentno in completed_students : continue
        markdata = { "mark": str(module_df["Mark"].at[row.Index]), "EBN": module_df["EBN"].at[row.Index] }
        if studentno not in students.keys() :
           localdict = {} 
           localdict["studentno"] = int(studentno)
           localdict["name"] = module_df['Student name'].at[row.Index] 
           localdict[modulename] = markdata
           students[int(studentno)] = localdict
        elif modulename in students[int(studentno)].keys() :
           if students[int(studentno)][modulename]["EBN"]=="U" : students[int(studentno)][modulename] = markdata
           elif "resit_EBN" in students[int(studentno)][modulename].keys() and students[int(studentno)][modulename]["resit_EBN"][-1]=="U"  :
              students[int(studentno)][modulename]["resit_mark"][-1] = str(module_df["Mark"].at[row.Index])
              students[int(studentno)][modulename]["resit_EBN"][-1] =  module_df["EBN"].at[row.Index] 
           else : 
              different=False
              for key, data in students[int(studentno)][modulename].items() :
                  if key not in markdata.keys() : raise Exception("mismatched keys in dictionaries")
                  elif data!=markdata[key] :
                     try: 
                        ddd = float(data) - float(markdata[key])
                        if ddd!=0 and not math.isnan(data) and not math.isnan(markdata[key]) : different=True
                     except: 
                        if markdata[key] not in data : different=True
              if different and "resit_EBN" not in students[int(studentno)][modulename].keys() : 
                 students[int(studentno)][modulename]["resit_mark"] = [markdata["mark"]]
                 students[int(studentno)][modulename]["resit_EBN"] = [markdata["EBN"]]
              elif different : raise Exception("Found erroneous marks for student " + str(studentno) + " in module " + modulename )
        elif "degprogram" in students[int(studentno)].keys() : 
           print( "Found rogue module for student", students[int(studentno)] )

print("GOT ALL MODULE DATA")

# Find out what marks are missing
missing_modules = set()
for key, value in students.items() :
    if value["studentno"] in completed_students : continue
    l1cats, l2cats, l3cats, l1average, l2average, l3average, l2fails, l3fails, l3absm, nmiss = 0, 0, 0, 0, 0, 0, 0, 0, 0, 0  
    for mkey, mvalue in value.items() :
        if mkey not in handbook.keys() : continue 
        if mvalue["EBN"]=="U" : 
           missing_modules.add(mkey)
           nmiss = nmiss + handbook[mkey]["CATS"]
        elif "resit_EBN" in mvalue.keys() and mvalue["resit_EBN"][-1]=="U" : 
           missing_modules.add(mkey)
           nmiss = nmiss + handbook[mkey]["CATS"]
           print("Missing resit marks for student", key, "taking module", mkey )
        elif handbook[mkey]["level"]>2 : 
           l3cats = l3cats + handbook[mkey]["CATS"]
           if "resit_EBN" not in mvalue.keys() : 
               if mvalue["EBN"]=="P" or mvalue["EBN"]=="PH" : l3average = l3average + handbook[mkey]["CATS"]*float(mvalue["mark"])
               elif "ABSM" in mvalue["EBN"] : 
                  l3cats = l3cats - handbook[mkey]["CATS"]
                  l3absm = l3absm + handbook[mkey]["CATS"] 
               else : 
                  l3cats = l3cats - handbook[mkey]["CATS"]
                  l3fails = l3fails + handbook[mkey]["CATS"]
                  if "F" in mvalue["EBN"] : l3average = l3average + handbook[mkey]["CATS"]*float(mvalue["mark"]) 
           elif mvalue["resit_EBN"][-1]=="P" : l3average = l3average + handbook[mkey]["CATS"]*float(mvalue["resit_mark"][-1])
           elif mvalue["resit_EBN"][-1]=="PH" : l3average = l3average + handbook[mkey]["CATS"]*40 
           else : 
              l3cats = l3cats - handbook[mkey]["CATS"]
              l3fails = l3fails + handbook[mkey]["CATS"]
              l3average = l3average + handbook[mkey]["CATS"]*getMaximumMark( mvalue )
        elif mvalue["EBN"]=="P" or "PAL" in mvalue["EBN"] or "PAS" in mvalue["EBN"] :
           if handbook[mkey]["level"]==1 : 
              l1cats = l1cats + handbook[mkey]["CATS"] 
              l1average = l1average + handbook[mkey]["CATS"]*float(mvalue["mark"])
           elif handbook[mkey]["level"]==2 : 
              l2cats = l2cats + handbook[mkey]["CATS"]
              l2average = l2average + handbook[mkey]["CATS"]*float(mvalue["mark"])
           else : raise Exception("Should not be finding a level three module here")
        elif mvalue["EBN"]=="PH" :
           print( "Found honours resitricted pass on first attempt for student " + str(key) + " in module " + mkey ) 
           if handbook[mkey]["level"]==1 :
              l1cats = l1cats + handbook[mkey]["CATS"]
              l1average = l1average + handbook[mkey]["CATS"]*40
           elif handbook[mkey]["level"]==2 :
              l2cats = l2cats + handbook[mkey]["CATS"]
              l2average = l2average + handbook[mkey]["CATS"]*40
           else : raise Exception("Should not be finding a level three module here") 
        elif "resit_EBN" in mvalue.keys() and (mvalue["resit_EBN"][-1]=="PH" or mvalue["resit_EBN"][-1]=="P") :
           if handbook[mkey]["level"]==1 :
              l1cats = l1cats + handbook[mkey]["CATS"]
              if "P" in mvalue["resit_EBN"][-1]=="P" : l1average = l1average + handbook[mkey]["CATS"]*float(mvalue["resit_mark"][-1])
              elif "PH" in mvalue["resit_EBN"][-1]=="PH" : l1average = l1average + handbook[mkey]["CATS"]*40
           elif handbook[mkey]["level"]==2 :
              l2cats = l2cats + handbook[mkey]["CATS"]
              if "P" in mvalue["resit_EBN"][-1]=="P" : l2average = l2average + handbook[mkey]["CATS"]*float(mvalue["resit_mark"][-1])
              elif "PH" in mvalue["resit_EBN"][-1]=="PH" : l2average = l2average + handbook[mkey]["CATS"]*40 
           else : raise Exception("Resit mark for level three student " + key  + " ????" )
        elif handbook[mkey]["level"]==2 : 
           l2fails = l2fails + handbook[mkey]["CATS"]
           l2average = l2average + handbook[mkey]["CATS"]*getMaximumMark( mvalue ) 
        else : 
           print("DISCARDING RESULT FOR STUDENT", key, "IN MODULE", mkey )
    value["L1CATS"] = l1cats 
    value["L2CATS"] = l2cats
    value["L2FAILS"] = l2fails
    value["L3CATS"] = l3cats
    value["L3ABSM"] = l3absm
    value["L3FAILS"] = l3fails
    if l1cats>0 : value["L1AVERAGE"] = l1average / l1cats 
    else : value["L1AVERAGE"] = l1average
    if l2cats>0 : value["L2AVERAGE"] = l2average / (l2cats + l2fails)
    else : value["L2AVERAGE"] = l2average
    if l3cats>0 : value["L3AVERAGE"] = l3average / (l3cats + l3fails + l3absm)
    else : value["L3AVERAGE"] = l3average
    if l1cats==120 and (l2cats+l2fails)==120 :
       if l3cats>0 and (l3cats + l3fails + l3absm)!=120 : print("WRONG L3CATS TOTAL FOR STUDENT", key, l3cats, l3fails, l3absm, (l3cats + l3fails + l3absm) )
       value["missing"] = nmiss
       if value["degprogram"]=="BSc" : 
          value["finalmark"] = 0.1*value["L1AVERAGE"] + 0.3*value["L2AVERAGE"] + 0.6*value["L3AVERAGE"]
          if l2cats==120 and l3cats==120 : 
             if value["finalmark"]>=70 : value["class"] = "1st"
             elif value["finalmark"]>=60 : value["class"] = "2.i"
             elif value["finalmark"]>=50 : value["class"] = "2.ii"
             elif value["finalmark"]>=40 : value["class"] = "3rd"
             else : value["class"] = "none"
          elif (value["L2FAILS"]+value["L3FAILS"])<=40 :
             if value["finalmark"]>=70 : value["class"] = "1st"
             elif value["finalmark"]>=60 : value["class"] = "2.i"
             elif value["finalmark"]>=50 : value["class"] = "2.ii"
             elif value["finalmark"]>=40 : value["class"] = "3rd"
             else : value["class"] = "none"
          elif value["missing"]>0 : value["class"] = "missing"
          elif value["L3ABSM"]>0 : value["class"] = "ABSM"
          else : value["class"] = "unset"
          if value["finalmark"]>40 and value["missing"]==0 : 
             b10 = 10*np.floor( value["finalmark"]/10 )
             if float(value["startyear"])<2201 and value["finalmark"]>=value["finalmark"]-b10>=7 : value["predominance"] = check_predominance( b10+10, value, handbook )
             elif value["finalmark"]>=value["finalmark"]-b10>=9 : value["predominance"] = check_predominance( b10+10, value, handbook )
             else : value["predominance"] = "no"
          else : value["predominance"] = "no"
       elif value["degprogram"]=="MSci" or value["degprogram"]=="MMath" or value["degprogram"]=="MPhys" : 
          value["finalmark"] = ( 5*value["L1AVERAGE"] + 15*value["L2AVERAGE"] + 30*value["L3AVERAGE"] ) / 50
          if value["finalmark"]>=55 : value["progress"] = "yes"
          elif value["L3ABSM"] : value["progress"] = "ABSM"
          else : value["progress"] = "no"
       else : raise Exception("Invalid degree program " + value["degprogram"] )


# Now split the data into our various dictionaries
bsc_dict, msc_dict, rand_dict, other_dict, placement_dict = {}, {}, {}, {}, {}
for key, value in students.items() :
    if value["L1CATS"]==120 and (value["L2CATS"]+value["L2FAILS"])==120 and value["L3CATS"]>0 : 
       if value["degprogram"]=="BSc" : bsc_dict[key] = getFinalModuleMarks( value, handbook )
       elif value["degprogram"]=="MSci" or value["degprogram"]=="MMath"or value["degprogram"]=="MPhys" : msc_dict[key] = getFinalModuleMarks( value, handbook )
       else : raise Exception("Invalid degree program " + value["degprogram"] )
    elif "degprogram" in value.keys() :
       if value["L3CATS"]==0 : 
          placement_dict[key] = getFinalModuleMarks( value, handbook )
       else :
          value["finalmark"] = "?"
          rand_dict[key] = getFinalModuleMarks( value, handbook )
    else : 
       other_dict[key] = getFinalModuleMarks( value, handbook )

print( "MISSING DATA ON MODULES", missing_modules )

with open('Results/student_marks.json', 'w' ) as fp : 
     json.dump( students, fp, indent=4 )

output_bsc = pd.DataFrame(bsc_dict).T
output_bsc.to_excel("Results/BSC_MARKS.xlsx" )

# Output MSci marks
output_msci = pd.DataFrame(msc_dict).T
output_msci.to_excel("Results/MSCI_MARKS.xlsx")

# Output L3 students who are probably on a placement
output_placement = pd.DataFrame(placement_dict).T
output_placement.to_excel("Results/PLACEMENT_MARKS.xlsx")

# Output L3 students who have done something weird
output_rand = pd.DataFrame(rand_dict).T
output_rand.to_excel("Results/MISC_MARKS.xlsx")

# Output L4 students and placement students who I don't care about
output_other = pd.DataFrame(other_dict).T
output_other.to_excel("Results/OTHER_MARKS.xlsx")
