from cmath import nan
import json
import os
import sys
from xml.dom import minidom
import pandas as pd
import xml.etree.ElementTree as ET
import uuid
import logging


xl_File = sys.argv[1]
coveragefile = os.getcwd()+"/sample/CP_Coverages.csv"
# print(xl_File)
OU = sys.argv[2]
formsFile = os.getcwd()+"/Forms.xlsx"
form = pd.read_excel(formsFile, sheet_name='Sheet1')
# df = pd.read_excel(xl_File, sheet_name='Sheet1')
df = pd.read_excel(xl_File, sheet_name='Sheet1')

coverage_name = pd.read_csv(coveragefile, encoding='latin-1')
json_file_path = os.getcwd()+"\line.json"
# print(json_file_path)
with open(json_file_path, 'r') as json_file:
    data = json.load(json_file)

logging.basicConfig(filename='Error.log', level=logging.DEBUG)
# Create a formatter and set it for the file handler
# file_handler = logging.FileHandler('custom_log_file.log')
# file_handler.setLevel(logging.DEBUG) 
# formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
# file_handler.setFormatter(formatter)

line_dict = data[0]
underwriting_comp = data[1]
avai_Jiras = []
All_tranaction = ['Renewal', 'Submission', 'PolicyChange','Rewrite','RewriteNewAccount','Reinstatement']

off = []
for i,co in df.iterrows():
 
   productcode =[]
   variable_fields = []
   prod = line_dict[co['Product'].strip()]
   if(prod in ['EPLine','DOLine','EOLine','FIDLine','GTCLine']):
       productcode += ['BHCPackage']
   elif prod == 'GL7Line':
       productcode += ['CPP7CommercialPackage','GL7GeneralLiability']
   elif prod == 'CP7Line' and OU == 'ASP':
       productcode += (['ESPackage'])
   elif prod == 'CP7Line':
       productcode += (['CPP7CommercialPackage','CP7CommercialProperty'])
   elif prod == 'WRBGeneralLiabilityLine' and OU != 'ASP':
       productcode += (['WRBGeneralLiability'])
   elif prod == 'WRBGeneralLiabilityLine' and OU == 'ASP':
       productcode += (['ESPackage','WRBGeneralLiability'])
   elif prod == 'WXSLine' and OU == 'ASP':
       productcode += (['ESPackage','WXSExcessLiability'])
   elif prod == 'WCULine':
       productcode += (['CPP7CommercialPackage'])
   elif prod == 'WCMLine':
      productcode += (['WCMWorkersComp'])
   elif prod == 'CA7Line':
      productcode += (['CA7CommAuto'])
   transaction = []
#    set(productcode)
   state = []
   Form_publicID = 'PCForm'+':'+str(uuid.uuid4())
   Form_N = str(co['FormNumber']).strip()
   Edition_N = str(co['Edition']).strip()
   # Description_ = co['Description']
   if "," in co['Transaction Type']:
       transaction = co['Transaction Type'].replace(" ", "").split(',')
   elif co['Transaction Type'] == 'All Except Cancellation':
       transaction = All_tranaction
   else:
       transaction.append(co['Transaction Type'])


   if  isinstance(co['State'], str) and "," in co['State']:
       state = co['State'].replace(" ", "").split(',')
   elif isinstance(co['State'], str):
       state.append(co['State'])
 
#    no_state = str(co['NoState'])
   no_State = []
   coveragecode = ''
   if isinstance(co['NoState'], str) and  "," in co['NoState']:
       no_State = co['NoState'].replace(" ", "").split(',')
   elif isinstance(co['NoState'], str) and co['NoState'] != 'N/A':
       no_State.append(co['NoState'])
   # coverage = co['CoverageCode']
   
   #Mapping values to FormPattern
   FormPattern = ET.Element('FormPattern',{'public-id':Form_publicID})

   ClausePatternCode = ET.SubElement(FormPattern, 'ClausePatternCode')
   for s,cov in coverage_name.iterrows():      
      # if Form_N == 'CA 01 06' and cov['CoverageName']=='Maryland Changes - Collision Coverage In Mexico':
      #   print(str(cov['LOB']).strip())
      #   print(str(cov['CoverageName']).strip())
      #   print(str(co['Description']).strip())
      if str(cov['LOB']).strip() == prod and str(cov['CoverageName']).lower().replace(" ", "").replace("-","").replace("–","") == str(co['Description']).lower().replace(" ", "").replace("-","").replace("–",""):
         coveragecode = cov['CoverageCode']
         ClausePatternCode.text = coveragecode
         if(cov['VariableName'] != 'Manual Premium' and cov['VariableName'] != 'NA' and cov['VariableName'] != 'nan'):
            variable_fields.append(str(cov['VariableCode']))
   Code = ET.SubElement(FormPattern, 'Code')
   codeCheck = (Form_N+ Edition_N).replace(' ','')
#    print(Code.text)
   for i,fm in form.iterrows():
      if (str(fm['Code']).strip().split('_')[0] == codeCheck.strip() or str(fm['Code']).strip().split('_')[0]+'_GL' == codeCheck.strip()) and str(fm['LOB']).strip() == prod.strip():
         Code.text = "dummy"
         logging.warning(f'{codeCheck.strip()} is Already Available')
      elif(str(fm['Code']).strip() == codeCheck.strip()):
        #  print(codeCheck, "idsjsndfjnjsnfdjknskdfnkskdfnksnkf")       
         # Code.text = codeCheck+'_ASP'
         Code.text = "dummy"
         break
   if Code.text == "dummy":
      continue
   if Code.text == None:
         Code.text = codeCheck
         
   CovSysTableValueExistsOnAll = ET.SubElement(FormPattern, 'CovSysTableValueExistsOnAll')
   CovSysTableValueExistsOnAll.text = 'false'


   CoverableSysTable = ET.SubElement(FormPattern, 'CoverableSysTable')
   CoverableType = ET.SubElement(FormPattern, 'CoverableType')
   CoverableTypeKey = ET.SubElement(FormPattern, 'CoverableTypeKey')
   CoverableTypeKeyExistsOnAll = ET.SubElement(FormPattern,'CoverableTypeKeyExistsOnAll')
   CoverableTypeKeys = ET.SubElement(FormPattern, 'CoverableTypeKeys')
   CoverableTypeList = ET.SubElement(FormPattern, 'CoverableTypeList')

   Description = ET.SubElement(FormPattern, 'Description')
   Description.text =  co['Description'].strip()

   Description_L10N_ARRAY = ET.SubElement(FormPattern, 'Description_L10N_ARRAY')
   Edition = ET.SubElement(FormPattern, 'Edition')
   Edition.text = Edition_N

   EndorsementNumbered = ET.SubElement(FormPattern, 'EndorsementNumbered')
   EndorsementNumbered.text = 'false'

   FormClassification = ET.SubElement(FormPattern, 'FormClassification')
   FormNumber = ET.SubElement(FormPattern, 'FormNumber')
   if OU in ['ADM','BHC','ASP']:
       FormNumber.text = Form_N.replace(' ','')
   else:
       FormNumber.text = Form_N

   FormPatternAdditionalInsuredTypes = ET.SubElement(FormPattern, 'FormPatternAdditionalInsuredTypes')
   FormPatternAdditionalInterestTypes = ET.SubElement(FormPattern, 'FormPatternAdditionalInterestTypes')
   FormPatternClauseCodes = ET.SubElement(FormPattern, 'FormPatternClauseCodes')
   FormPatternCovTerms = ET.SubElement(FormPattern, 'FormPatternCovTerms')

   if len(variable_fields) >= 1:
     for ele in variable_fields:
        FormPatternCovTerm = ET.SubElement(FormPatternCovTerms,'FormPatternCovTerm',{'public-id':str(uuid.uuid4())})
        CovTermPatternCode = ET.SubElement(FormPatternCovTerm,'CovTermPatternCode')
        CovTermPatternCode.text = ele
        CovFormPattern = ET.SubElement(FormPatternCovTerm,'FormPattern',{'public-id':Form_publicID})
        SelectedCovTermValues = ET.SubElement(FormPatternCovTerm,'SelectedCovTermValues')

   FormPatternCoverableProperties = ET.SubElement(FormPattern, 'FormPatternCoverableProperties')
   FormPatternCoveragePartTypes = ET.SubElement(FormPattern, 'FormPatternCoveragePartTypes')

   #FormPattern Job
   FormPatternJobs = ET.SubElement(FormPattern, 'FormPatternJobs')
   
   for types in transaction:
         FormPattenJob = ET.SubElement(FormPatternJobs,'FormPatternJob',{'public-id':str(uuid.uuid4())})
         JobFP = ET.SubElement(FormPattenJob,'FormPattern',{'public-id':Form_publicID})
         if types == 'Policychange' or types == 'PolicyChange':
            ET.SubElement(FormPattenJob,'JobType').text = 'PolicyChange'
         else:
            ET.SubElement(FormPattenJob,'JobType').text = types
   
   #FormPatternOU_Ext
   Offerings_Ext = ET.SubElement(FormPattern, 'Offerings_Ext')
   for offering in off:
      FormPatternOffering_Ext = ET.SubElement(Offerings_Ext,'FormPatternOffering_Ext',{'public-id':str(uuid.uuid4())})
      Avl = ET.SubElement(FormPatternOffering_Ext,'Availability')
      Avl.text = 'Unavailable'
      ouForm = ET.SubElement(FormPatternOffering_Ext,'FormPattern',{'public-id':Form_publicID})
      OfferingCode = ET.SubElement(FormPatternOffering_Ext,'OfferingCode')
      # if prod in ['EPLLine','DOLLine','EOLine','FIDLine',None]:
      OfferingCode.text = offering

   #Form Product
   FormPatternProducts = ET.SubElement(FormPattern, 'FormPatternProducts')
   for j in productcode:
     Form_prd = ET.SubElement(FormPatternProducts, 'FormPatternProduct',{'public-id':str(uuid.uuid4())})
     Fromprod = ET.SubElement(Form_prd,'FormPattern',{'public-id':Form_publicID})
     prodCode = ET.SubElement(Form_prd, 'ProductCode')
     prodCode.text = j   

   

   GenericInferenceClass = ET.SubElement(FormPattern, 'GenericInferenceClass')
   if coveragecode != '':
      GenericInferenceClass.text = 'gw.forms.generic.GenericClauseSelectionForm'
   elif state == ['CW']:
      GenericInferenceClass.text = 'gw.forms.generic.GenericAlwaysAddedForm'
   else:
      GenericInferenceClass.text = 'com.wrberkley.form.GenericBaseStateApplicable'
   InferenceTime = ET.SubElement(FormPattern, 'InferenceTime')
   InferenceTime.text = 'quote'
   InternalGroupCode = ET.SubElement(FormPattern, 'InternalGroupCode')
   InternalReissueOnChange = ET.SubElement(FormPattern, 'InternalReissueOnChange')
   print(transaction)
   if  'PolicyChange' in transaction or 'Policychange' in transaction:
      InternalReissueOnChange.text = 'true'
   else:
      InternalReissueOnChange.text = 'false'
   InternalRemovalEndorsement = ET.SubElement(FormPattern, 'InternalRemovalEndorsement')

   Lookups = ET.SubElement(FormPattern, 'Lookups')
   print(state)
   if state is not None:
     for lk in state:
       if isinstance(co['Company'],str):
         for cm in co['Company'].split(','):
            FpLookup = ET.SubElement(Lookups,'FormPatternLookup',{'public-id':str(uuid.uuid4())})
            AvlF = ET.SubElement(FpLookup,'Availability')
            AvlF.text = 'Available'
            Efd = ET.SubElement(FpLookup,'EndEffectiveDate')
            FpLP = ET.SubElement(FpLookup,'FormPattern',{'public-id':Form_publicID})
            Sedate = ET.SubElement(FpLookup,'StartEffectiveDate')
            Sedate.text = '2023-01-01 00:00:00.000'
            jurisdiction = ET.SubElement(FpLookup,'Jurisdiction')
            if lk != 'CW':
              jurisdiction.text = lk
            uwcomp = ET.SubElement(FpLookup,'UWCompanyCode')
            print(cm + "---------------------")
            # if(cm.strip() in underwriting_comp):
            #    uwcomp.text =  underwriting_comp[cm.strip()]
       else:
            FpLookup = ET.SubElement(Lookups,'FormPatternLookup',{'public-id':str(uuid.uuid4())})
            AvlF = ET.SubElement(FpLookup,'Availability')
            AvlF.text = 'Available'
            Efd = ET.SubElement(FpLookup,'EndEffectiveDate')
            FpLP = ET.SubElement(FpLookup,'FormPattern',{'public-id':Form_publicID})
            Sedate = ET.SubElement(FpLookup,'StartEffectiveDate')
            Sedate.text = '2023-01-01 00:00:00.000'
            jurisdiction = ET.SubElement(FpLookup,'Jurisdiction')
            if lk != 'CW':
              jurisdiction.text = lk
            # if lk != 'CW':
            uwcomp = ET.SubElement(FpLookup,'UWCompanyCode')
   else:
       FpLookup = ET.SubElement(Lookups,'FormPatternLookup',{'public-id':str(uuid.uuid4())})
       AvlF = ET.SubElement(FpLookup,'Availability')
       AvlF.text = 'Available'
       Efd = ET.SubElement(FpLookup,'EndEffectiveDate')
       FpLP = ET.SubElement(FpLookup,'FormPattern',{'public-id':Form_publicID})
       Sedate = ET.SubElement(FpLookup,'StartEffectiveDate')
       Sedate.text = '2023-01-01 00:00:00.000'
       jurisdiction = ET.SubElement(FpLookup,'Jurisdiction')
       uwcomp = ET.SubElement(FpLookup,'UWCompanyCode')
    #    if isinstance(co['Company'],str) and ',' not in co['Company']:
    #      uwcomp.text =  underwriting_comp[co['Company']]

   if no_State is not None:
     for lj in no_State:
       if isinstance(co['Company'],str):
         for cp in co['Company'].split(','):
            FpLookup = ET.SubElement(Lookups,'FormPatternLookup',{'public-id':str(uuid.uuid4())})
            AvlF = ET.SubElement(FpLookup,'Availability')
            AvlF.text = 'Unavailable'
            Efd = ET.SubElement(FpLookup,'EndEffectiveDate')
            FpLP = ET.SubElement(FpLookup,'FormPattern',{'public-id':Form_publicID})
            Sedate = ET.SubElement(FpLookup,'StartEffectiveDate')
            Sedate.text = '2023-01-01 00:00:00.000'
            jurisdiction = ET.SubElement(FpLookup,'Jurisdiction')
            if lj != 'CW':
              jurisdiction.text = lj
            uwcomp = ET.SubElement(FpLookup,'UWCompanyCode')
            # if(cm.strip() in underwriting_comp):
               # uwcomp.text =  underwriting_comp[cm.strip()]
       else:
            FpLookup = ET.SubElement(Lookups,'FormPatternLookup',{'public-id':str(uuid.uuid4())})
            AvlF = ET.SubElement(FpLookup,'Availability')
            AvlF.text = 'Unavailable'
            Efd = ET.SubElement(FpLookup,'EndEffectiveDate')
            FpLP = ET.SubElement(FpLookup,'FormPattern',{'public-id':Form_publicID})
            Sedate = ET.SubElement(FpLookup,'StartEffectiveDate')
            Sedate.text = '2023-01-01 00:00:00.000'
            jurisdiction = ET.SubElement(FpLookup,'Jurisdiction')
            if lj != 'CW':
              jurisdiction.text = lj
            uwcomp = ET.SubElement(FpLookup,'UWCompanyCode')  
   else:       
       FpLookup = ET.SubElement(Lookups,'FormPatternLookup',{'public-id':str(uuid.uuid4())})
       AvlF = ET.SubElement(FpLookup,'Availability')
       AvlF.text = 'Unavailable'
       Efd = ET.SubElement(FpLookup,'EndEffectiveDate')
       FpLP = ET.SubElement(FpLookup,'FormPattern',{'public-id':Form_publicID})
       Sedate = ET.SubElement(FpLookup,'StartEffectiveDate')
       Sedate.text = '2023-01-01 00:00:00.000'
       jurisdiction = ET.SubElement(FpLookup,'Jurisdiction')
       uwcomp = ET.SubElement(FpLookup,'UWCompanyCode')
    #    if isinstance(co['Company'],str) and ',' not in co['Company']:
    #      uwcomp.text =  underwriting_comp[co['Company']] 
   not_line = ['CPP7CommercialPackage']
   PatternExistsOnAll = ET.SubElement(FormPattern, 'PatternExistsOnAll')
   PatternExistsOnAll.text = 'false'
   PolicyLinePatternCode = ET.SubElement(FormPattern, 'PolicyLinePatternCode')
   if prod != 'GTCLine' and sum(1 for value in prod if value not in not_line) >= 1:
      PolicyLinePatternCode.text = prod
   Priority = ET.SubElement(FormPattern, 'Priority')
   Priority.text = '-1'
   RefCode = ET.SubElement(FormPattern, 'RefCode')
   RefCode.text = Form_N.replace(' ','') + '_' + Edition_N.replace(' ','')
   RemovalEndorsementFormNumber = ET.SubElement(FormPattern, 'RemovalEndorsementFormNumber')
   SequenceType_Ext = ET.SubElement(FormPattern, 'SequenceType_Ext')
   UseInsteadOfCode = ET.SubElement(FormPattern, 'UseInsteadOfCode')
   

   tree = ET.ElementTree(FormPattern)
   xml_str = ET.tostring(FormPattern,encoding='utf-8').decode('utf-8')

   dom = minidom.parseString(xml_str)
   pretty_xml_str = dom.toprettyxml(indent=" ")
   pretty_xml_str = pretty_xml_str.replace('<?xml version="1.0" ?>\n', '')


   with open(f"{prod}FormPatterns.xml", 'a') as file:
       print(f"{codeCheck} is created......")
       print(str(co['Jira']))
       avai_Jiras.append(str(co['Jira']))
       file.write(pretty_xml_str)
print(avai_Jiras)


