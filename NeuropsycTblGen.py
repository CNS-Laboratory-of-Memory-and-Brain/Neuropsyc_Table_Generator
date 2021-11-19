'''
Neuropsych Table Generator
Author: Zimu Li
Organization: Smith Memory Lab
@param xlsx files in format that allign with NP_100819.xlsx
'''
import pandas as pd


def main():
	inputname = input("Enter input file path : ")
	outputname = input("Enter outputfile path : ")

	calibration = input("is calibration needed?(Y/N)")

	tbl = pd.read_excel(inputname,index_col=0)
	xls = pd.ExcelFile(inputname)
	tblPVal = pd.read_excel(xls,'ANCOVA')
	ConStartRow = 0
	MCIStartRow = 0
	numOfEntries = 83
	indecies = tbl.index.tolist()
	for i in range(len(indecies)):
		print(indecies[i])
		if indecies[i] == 'Group = 1.00 Normal' and indecies[i+5] == 'Age':
			ConStartRow = i+5
		if indecies[i] == 'Group = 2.00 MCI' and indecies[i+5] == 'Age':
			MCIStartRow = i+5

	if calibration == 'Y':
		print(tbl.head(300))
		ConStartRow = input("Please enter the index number where control stats starts(usually the index of \"Age\"): ")
		MCIStartRow = input("Please enter the index number where MCI stats starts(usually the index of \"Age\"): ")
		numOfEntries = input("Please enter the number of entries per group: ")

	tblControl = tbl.iloc[ConStartRow:ConStartRow + numOfEntries,[3,5]]
	tblMCI = tbl.iloc[MCIStartRow:MCIStartRow + numOfEntries,[3,5]]

	if calibration == 'Y':
		print('please check the tables')
		print(tblControl.head(numOfEntries))
		print(tblMCI.head(numOfEntries))

	tblControl['CON_Mean(sd)'] = tblControl.apply(lambda x: str(round(x['Unnamed: 4'],1)) + ' (' + str(round(x['Unnamed: 6'],1)) + ')', axis = 1)
	tblMCI['MCI_Mean(sd)'] = tblMCI.apply(lambda x: str(round(x['Unnamed: 4'],1)) + ' (' + str(round(x['Unnamed: 6'],1)) + ')', axis = 1)

	df1 = tblControl.merge(tblMCI, left_index=True, right_index=True)

	df1 = df1.drop(columns=['Unnamed: 4_x','Unnamed: 6_x','Unnamed: 4_y','Unnamed: 6_y'],axis=1)

	calibrationSpecify = input('Do you need to re-specify the tests needed?(Y/N) : ')

	desiredData = ['Age','Education','Total_Visits','Mini_Mental','CVLT_ToT_Recall_R','CVLT_ToT_Recog_R', \
	'CVLT_LDF_Recall_R','WMS_LM_I_Recall_R','WMS_LM_D_Recall_R','WMS_VR_I_Recall_R','WMS_VR_D_Recall_R' \
	, 'DKEFS_CategoryFluency_R','DKEFS_LetterFluency_R','MINT_R','Multilingual_AE_TokenT_R','WRAT_BW_Reading_R' \
	, 'WAIS_DigitSpan_Backward_R', 'DKEFS_Fluency_Switching_R', 'DKEFS_Trail_Making_C4_R', 'WCS_Categories_R', \
	'WCS_Perseverative_Errors_R', 'DKEFS_Trail_Making_C2_R', 'WAIS_DigitSpan_Seq_R', 'WAIS_DigitSpan_Forward_R', \
	'Digit_Vigilance_Total_Time_R', 'Digit_Vigilance_Errors_R','Clock_Drawing_Command_R','Clock_Drawing_Copy_R', \
	'Overlapping_Pentagons_R', 'WASI_II_BlockDesign_R', 'WMS_IV_VisualRepro_Copy_R', 'ILS_Health_Safety_R', \
	'ILScale_Managing_Money_R', 'FAQ_R', 'Geriatric_Depression_Inventory']

	if calibrationSpecify == 'Y':
		desiredData = []
		numspecs = input('Please enter the total number of tests per group: ')
		for i in range(numspecs):
			testName = input('Please enter the name of the test #' + str(i))
			desiredData.append(testName)
		print('We will collect mean and sd from the following test')
		print(desiredData)

	df1Desired = df1.loc[desiredData,:]

	col = tblPVal['Unnamed: 1'].tolist()

	pval = [float('nan'),float('nan'),float('nan')]
	pv = tblPVal['Unnamed: 5'].tolist()
	for i in range(len(col)):
	    if col[i] in desiredData and col[i+1] == 'Type III Sum of Squares':
	        pval.append(round(pv[i+6],3))

	df1Desired['p-value'] = pval


	renameIndecies = {"Mini_Mental": "Mini-Mental State Exam", "CVLT_ToT_Recall_R":"CVLT-II: Trials 1-5 Total Recall"\
		, "CVLT_ToT_Recog_R":"CVLT-II: Recognition (d\')", "CVLT_LDF_Recall_R'":"CVLT-II: Long-Delay Free Recall " \
		, "WMS_LM_I_Recall_R":"WMS-IV Logical Memory: Immediate Recall", "WMS_LM_D_Recall_R":"WMS-IV Logical Memory: Delayed Recall" \
		, "WMS_VR_I_Recall_R":"WMS IV Visual Repro: Immediate Recall", "WMS_VR_D_Recall_R":"WMS IV Visual Repro: Delayed Recall" \
		, "DKEFS_CategoryFluency_R":"DKEFS: Category Fluency (Animals, Boys Names)", "DKEFS_LetterFluency_R":"DKEFS: Letter Fluency (F, A, S words)" \
		, "MINT_R":"Multilingual Naming Test (MINT)", "Multilingual_AE_TokenT_R":"Multilingual Aphasia Exam: Token Test" \
		, "WRAT_BW_Reading_R":"WRAT4 Blue word reading test", "WAIS_DigitSpan_Backward_R":"WAIS-IV Digit Span (Backward)" \
		, "DKEFS_Fluency_Switching_R":"D-KEFS Fluency Switching (Fruits/Furniture)", "DKEFS_Trail_Making_C4_R":"D-KEFS Trail Making Cond 4 (equiv to Trail B)" \
		, "WCS_Categories_R":"Wisconsin Card Sorting: Categories", "WCS_Perseverative_Errors_R":"Wisconsin Card Sorting: Perseverative Errors"\
		, "DKEFS_Trail_Making_C2_R":"D-KEFS Trail Making Cond 2 (equiv to Trail A)", "WAIS_DigitSpan_Seq_R":"WAIS-IV Digit Span (Sequencing)"\
		, "Digit_Vigilance_Total_Time_R":"Digit Vigilance Test: Total Time", "Digit_Vigilance_Errors_R":"Digit Vigilance Test: Errors"\
		, "Clock_Drawing_Command_R":"Clock Drawing Test Command", "Clock_Drawing_Copy_R":"Clock Drawing Test Copy" \
		, "Overlapping_Pentagons_R":"Overlapping Pentagons", "WASI_II_BlockDesign_R":"WASI-II Block Design"\
		, "WMS_IV_VisualRepro_Copy_R":"WMS-IV Visual Repro: Copy", "ILS_Health_Safety_R":"Independent Living Scale: Health and Safety"\
		, "FAQ_R":"Functional Activities Questionaire (FAQ)", "Geriatric_Depression_Inventory":"Geriatric Depression Inventory",\
		"WAIS_DigitSpan_Forward_R" : "WAIS-IV Digit Span (Forward)",\
		"ILScale_Managing_Money_R" : "Independent Living Scale: Managing Money"}

	if calibrationSpecify == 'Y':
		print('please re-specify renames of the indices')
		renameIndecies = {}
		for i in desiredData:
			renamed = input('What should test ' + i + ' be rename to? : ')
			renameIndecies[i] = renamed

	df1Desired.rename(index = renameIndecies,inplace = True) 

	if calibrationSpecify != 'Y':
		dfDemographic = df1Desired.iloc[:3,:]
		dfGlobalCognition = df1Desired.iloc[3:4,:]
		dfEpisodicMemory = df1Desired.iloc[4:11,:]
		dfSemanticMemory = df1Desired.iloc[11:16,:]
		dfExecutiveFunctions = df1Desired.iloc[16:21,:]
		dfAttention = df1Desired.iloc[21:26,:]
		dfVisuospatial = df1Desired.iloc[26:31,:]
		dfFunctionalAbilities = df1Desired.iloc[31:34,:]
		dfEmotional = df1Desired.iloc[34:,:]

		def addTitle(df, title):
		    new_row = pd.DataFrame({'CON_Mean(sd)':float('nan'),'MCI_Mean(sd)':float('nan'),'p-value':float('nan')}, index =[title])
		    df = pd.concat([new_row, df[:]]).reset_index(drop = False) 
		    return df

		dff1 = addTitle(dfDemographic,'Demographic:')
		dff2 = addTitle(dfGlobalCognition, 'Global Cognitions:')
		dff3 = addTitle(dfEpisodicMemory, 'Episodic Memory:')
		dff4 = addTitle(dfSemanticMemory, 'Semantic Memory/Language:')
		dff5 = addTitle(dfExecutiveFunctions, 'Executive Functions:')
		dff6 = addTitle(dfAttention, 'Attention/Processing Speed:')
		dff7 = addTitle(dfVisuospatial, 'Visuospatial Functions:')
		dff8 = addTitle(dfFunctionalAbilities, 'Functional Abilities:')
		dff9 = addTitle(dfEmotional, 'Emotional/Health:')

		resultdf = pd.concat([dff1,dff2,dff3,dff4,dff5,dff6,dff7,dff8,dff9])

		print(resultdf.head(100))

		resultdf.to_excel(outputname,index= False)
	else:
		df1Desired.to_excel(outputname,index = True)
		
	print('Neuropsych Table successfully generated!!')


main()
