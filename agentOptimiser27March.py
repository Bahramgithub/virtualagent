#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue Mar  5 16:28:59 2019

@author: bahram.vazir.nezhad@accenture.com
"""

import os, re, xlwt#, xlrd
import dialogflow_v2 as dialogflow
import pandas as pd
import numpy as np
from intentManage import intentManage as intentManager
from time import sleep

## authenticate to the API using your service account credentials
minTrainSizeList = [8, 13] # [5, 8, 13] #
confThr = [0.3, 0.4, 0.5, 0.6]
inScopeIntents = ["BalanceCheck", "BillRequest", "BillPay", "DirectDebitChange", "ContractExpiryRequest", "AgentHandover", "PaymentExtend", "OrderEnquire", "RoamingInformationRequest"]# , "SimActivate", "InternetAccess", "AccountDetailsChange" 
#inScopeIntentsLive = ["BalanceCheck", "BillRequest", "BillPay", "DirectDebitChange", "ContractExpiryRequest", "AgentHandover", "PaymentExtend"] # "AgentHandover", "DirectDebitChange", "Farewell", "SimActivate", "InternetAccess", "AccountDetailsChange" 
#inScopeIntents = inScopeIntentsLive
inScopeIntentsWeights = {"BalanceCheck" : 8, "BillRequest" : 10,
                         "BillPay" : 7, "ContractExpiryRequest" : 43,
                         "DirectDebitChange" : 11, "AgentHandover" : 59,
                         "PaymentExtend" : 24, "OrderEnquire" : 37,
                         "RoamingInformationRequest" : 19}

totalWeight = 0
for intent in inScopeIntentsWeights:
    totalWeight += inScopeIntentsWeights [intent]
bestPerformance = 0 #bestAccuracy = 0
#bestModifiedAccuracy = 0

data = pd.ExcelFile ("RoutingBotMaster_20191023.xlsx")
intentDicTrain = {}
inputFrame1 = data.parse ("MainUtterances")

intentList1 = list (set (inputFrame1 ["Final Intent"]))
for intent in intentList1:
    intentDicTrain [intent] = []

for row in range (len (inputFrame1)):
    if (inputFrame1 ["Fixed Test Train"][row] == "TRAIN"):# and inputFrame1 ["Source"][row] == "Live"):
        intentDicTrain [inputFrame1 ["Final Intent"][row]].append (inputFrame1 ["Redacted Utterance"][row])

#try:
#    del intentDicTrain ["IncompleteIntent"]
#except:
#    pass
#try:
#    del intentDicTrain ["NS"]
#except:
#    pass
#try:
#    del intentDicTrain ["Startover"]
#except:
#    pass
#try:
#    del intentDicTrain ["Farewell"]
#except:
#    pass

##########

for minTrainSize in minTrainSizeList:
    
    os.environ ["GOOGLE_APPLICATION_CREDENTIALS"] = "testagent-2699e-7cddbe973592.json"
    session_client = dialogflow.SessionsClient ()

    projectId = "testagent-2699e"

    trainableIntents = []
    for intent in intentDicTrain:
        if (len (intentDicTrain [intent]) >= minTrainSize):
            trainableIntents.append (intent)
    
    intentIds = intentManager.list_intents (projectId)
    for intentId in intentIds:
        index = [(i.start(), i.end()) for i in re.finditer ("/", intentId)][-1][1]
        Id = intentId [index : len (intentId)]
        intentManager.delete_intent (projectId, Id)
    
    trainableModel = 0
    for intent in trainableIntents:
        intentManager.create_intent (projectId, intent, intentDicTrain [intent], [intent])
        trainableModel += 1
    
    # results with various thresholds #####################################
    # move intent detection out of the loop 
#    for utterance in intentDicTest [intent]:
#                if len (utterance) > 255:
#                    utterance = utterance [0 : 255]
#                sleep (0.3)
#                [identifiedIntent, conf, response] = intentManager.detect_intent_texts (projectId, "abc123", [utterance], "en")

    for thr in confThr:
        
        fileName = str (minTrainSize) + "-" + str (thr) + "-23Oct-V0.xls"   
        outputWorkbook = xlwt.Workbook (encoding = "utf-8")
        resultSheet = outputWorkbook.add_sheet ("Stat ", cell_overwrite_ok = True)
        confusionSheet = outputWorkbook.add_sheet ("Confusion matrix ", cell_overwrite_ok = True)
        utteranceSheet = outputWorkbook.add_sheet ("Utterances ", cell_overwrite_ok = True)

        session_client = dialogflow.SessionsClient ()
        
        intentDicTest = {}
        statDic = {}
        intentDicTest ["HO"] = []
        statDic ["HO"] = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0] # TP, FN, FP, TN, Precision, Recall, F1, Accuracy, count, confidence
        intentDicTest ["FB"] = []
        statDic ["FB"] = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0] # TP, FN, FP, TN, Precision, Recall, F1, Accuracy, count, confidence
        
        for intent in intentList1:
            if (intent in trainableIntents):
                intentDicTest [intent] = []
                statDic [intent] = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0] # TP, FN, FP, TN, Precision, Recall, F1, Accuracy, count, confidence
                
        c = 0
        for row in range (len (inputFrame1)):
            if (inputFrame1 ["Final Intent"][row] == "IncompleteIntent"):
                continue
            if (inputFrame1 ["Fixed Test Train"][row] == "TEST"): # and inputFrame1 ["Source"][row] == "Live"):# and (inputFrame1 ["Final Intent"][row] in intentDicTrain.keys())):
                if (inputFrame1 ["Final Intent"][row] in inScopeIntents): # and inputFrame1 ["Final Intent"][row] in trainableIntents): # intentDicTrain.keys ()):
                    intentDicTest [inputFrame1 ["Final Intent"][row]].append (inputFrame1 ["Redacted Utterance"][row])
                elif ((inputFrame1 ["Final Intent"][row] in trainableIntents) and (inputFrame1 ["Final Intent"][row] not in inScopeIntents)):##
                    intentDicTest ["HO"].append (inputFrame1 ["Redacted Utterance"][row])
                elif (inputFrame1 ["Final Intent"][row] == "NS"):
                    continue
#                    intentDicTest ["FB"].append (inputFrame1 ["Optimised Utterance for Training"][row])
                elif (inputFrame1 ["Final Intent"][row] not in trainableIntents): # and (inputFrame1 ["Final Intent"][row] not in inScopeIntents)): 
                    intentDicTest ["FB"].append (inputFrame1 ["Redacted Utterance"][row])##
                c += 1
##
#        try:
#            del intentDicTest ["IncompleteIntent"]
#        except:
#            pass
                
#        intentClasses = trainableIntents.copy ()
#        for intent in intentDicTest:
#            if not intentDicTest [intent]:
#                del intentDicTest [intent]
#        intentClasses = []
#        for intent in intentDicTest:
#            if intentDicTest [intent]:
#                intentClasses.append (intent)
#        for intent in intentClasses:
#            if not intentDicTest [intent]:
#                del intentDicTest [intent]
        intentClasses = inScopeIntents.copy ()
        intentClasses.append ("HO")
        intentClasses.append ("FB")
        dimension = (len (intentClasses), len (intentClasses))
        matrix = np.zeros (dimension) # confusion matrix
        cnt = 0
        utteranceRows = []
        cntHO = 0
        cntFB = 0
        costFB = 0
        for intent in intentDicTest:
            cnt1 = 0
            accConf = 0
            intentMapped = intent
            for utterance in intentDicTest [intent]:
                if len (utterance) > 255:
                    utterance = utterance [0 : 255]
                sleep (0.3)
                [identifiedIntent, conf, response] = intentManager.detect_intent_texts (projectId, "abc123", [utterance], "en")
                print (utterance)
                if conf < thr:
                    identifiedIntent = "FB" # no model and Out of Scope
                elif ((identifiedIntent in intentDicTrain.keys ()) and (identifiedIntent not in inScopeIntents)):
                    identifiedIntent = "HO"

                if (intent == "HO"):# or (intent in inScopeIntents)):
                    intentMapped = "HO"
                    cntHO += 1
                elif intent not in trainableIntents:##
                    intentMapped = "FB"
                    cntFB += 1
#                elif ((intent in intentDicTrain.keys ()) and (intent not in inScopeIntents)):
#                    intentMapped = "HO"
                
                if identifiedIntent == intentMapped:
                    statDic [intentMapped][0] += 1 # TP += 1
                    matrix [intentClasses.index (intentMapped), intentClasses.index (intentMapped)] += 1
                else:
                    statDic [intentMapped][1] += 1 # FN += 1
                    statDic [identifiedIntent][2] += 1 # FP += 1
                    matrix [intentClasses.index (intentMapped), intentClasses.index (identifiedIntent)] += 1
                    if identifiedIntent == "FB":
                        if (intentMapped in inScopeIntents and intentMapped != "AgentHandover"):
                            costFB += 1#0.05 #0.1 # 0.05 # 0.02
                        elif (intentMapped == "AgentHandover"):
                            costFB += 1#0.075 #0.15 # 0.075 # 0.03
                        else:
                            costFB += 1#0.025 #0.05 # 0.025 # 0.01
                cnt1 += 1
                accConf += conf
                cnt += 1
                utteranceRows.append ([utterance, intentMapped, identifiedIntent, conf, (intentMapped == identifiedIntent)])
#            if intentMapped in inScopeIntents:
            statDic [intentMapped][8] = cnt1
            statDic ["HO"][8] = cntHO
            statDic ["FB"][8] = cntFB
            try:
                statDic [intentMapped][9] = accConf / (statDic [intentMapped][0] + statDic [intentMapped][1])
            except:
                pass
        totalTP = 0  
        for intent in intentClasses:
            statDic [intent][3] = cnt - statDic [intent][0] - statDic [intent][1] - statDic [intent][2] # TN
            try:
                statDic [intent][4] = statDic [intent][0] / (statDic [intent][0] + statDic [intent][2]) # Precision
            except: 
                pass
            try:
                statDic [intent][5] = statDic [intent][0] / (statDic [intent][0] + statDic [intent][1]) # Recall
            except:
                pass
            try:
                statDic [intent][7] = (statDic [intent][0] + statDic [intent][3]) / (statDic [intent][0] + statDic [intent][1] + statDic [intent][2] + statDic [intent][3]) # Accuracy
            except:
                pass 
            try:
                statDic [intent][6] = (1.25 * statDic [intent][4] * statDic [intent][5]) / (0.25 * statDic [intent][4] + statDic [intent][5]) # F1
            except:
                pass
            totalTP += statDic [intent][0]
        
        rowCnt = 0
        resultSheet.write (rowCnt, 0, "Intent")
        resultSheet.write (rowCnt, 1, "TP")
        resultSheet.write (rowCnt, 2, "FN")
        resultSheet.write (rowCnt, 3, "FP")
        resultSheet.write (rowCnt, 4, "TN")
        resultSheet.write (rowCnt, 5, "Precision")
        resultSheet.write (rowCnt, 6, "Recall")
        resultSheet.write (rowCnt, 7, "F1")
        resultSheet.write (rowCnt, 8, "Accuracy")
        resultSheet.write (rowCnt, 9, "Count")
        resultSheet.write (rowCnt, 10, "Mean Confidence")
        resultSheet.write (rowCnt, 14, "WISF")#"Total Accuracy")
        resultSheet.write (rowCnt, 15, "Count FB")
        resultSheet.write (rowCnt, 16, "Min acceptable train")
        resultSheet.write (rowCnt, 17, "Threshold")
        resultSheet.write (rowCnt, 18, "Number of models")
        # change criteria also here for F1/P/R on in-scopes or all
        WISF1 = 0
        for intent in intentClasses:
            rowCnt += 1
            if intent in inScopeIntentsWeights.keys():
                aff = str (inScopeIntentsWeights [intent])
            else:
                aff = "NA"
            resultSheet.write (rowCnt, 0, intent + "-" + aff)
            resultSheet.write (rowCnt, 1, statDic [intent][0])
            resultSheet.write (rowCnt, 2, statDic [intent][1])
            resultSheet.write (rowCnt, 3, statDic [intent][2])
            resultSheet.write (rowCnt, 4, statDic [intent][3])
            resultSheet.write (rowCnt, 5, statDic [intent][4])
            resultSheet.write (rowCnt, 6, statDic [intent][5])
            resultSheet.write (rowCnt, 7, statDic [intent][6])
            resultSheet.write (rowCnt, 8, statDic [intent][7])
            resultSheet.write (rowCnt, 9, statDic [intent][8])
            resultSheet.write (rowCnt, 10, statDic [intent][9])
            if intent in inScopeIntents:
                weight = inScopeIntentsWeights [intent]
                WISF1 = WISF1 + (weight * statDic [intent][6] / totalWeight)
        performance = WISF1# - costFB
        totalAccuracy = totalTP / cnt
#        ISF1 = ISF1 / len (inScopeIntents)
        # change criteria here
        if (performance > bestPerformance and costFB < cnt * 0.15): # totalAccuracy >= bestAccuracy:
            bestPerformance = performance # bestAccuracy = totalAccuracy
            optimumParameters = [minTrainSize, thr]
        
#        inClassSet = set (intentClasses)
#        inScopeSet = set (inScopeIntents)
#        handoverSet = inClassSet.difference (inScopeSet)
#        handoverList = list (handoverSet.difference (set (["FB"])))
#        modifiedTotalTP = 0
#        for row in range (dimension [0]):
#            for col in range (dimension [1]):
#                if ((intentClasses [row] not in handoverList) and (intentClasses [col] not in handoverList)):
#                    if row == col:
#                        modifiedTotalTP = modifiedTotalTP + matrix [row, col]
#                elif ((intentClasses [row] in handoverList) and (intentClasses [col] in handoverList)):
#                    modifiedTotalTP = modifiedTotalTP + matrix [row, col]
#        modifiedTotalAccuracy = modifiedTotalTP / cnt
#        if modifiedTotalAccuracy > bestModifiedAccuracy:
#            bestModifiedAccuracy = modifiedTotalAccuracy
#            optimumParametersModified = [minTrainSize, thr]
        # REMOVE NS, INCOMPLETE, OTHERS FROM MASTERFILE AND USE WRONG FBS TO PENALISE THE MODEL        
        print (WISF1, thr, minTrainSize)#(totalAccuracy, modifiedTotalAccuracy, thr, minTrainSize)
        resultSheet.write (1, 14, WISF1)#(1, 14, totalAccuracy)
        resultSheet.write (1, 15, costFB)
        resultSheet.write (1, 16, minTrainSize)
        resultSheet.write (1, 17, thr)
        resultSheet.write (1, 18, trainableModel)
        for intent in intentClasses:
            confusionSheet.write (intentClasses.index (intent) + 1, 0, intent) 
            confusionSheet.write (0, intentClasses.index (intent) + 1, intent)
        for intent in intentClasses:
            for identifiedIntent in intentClasses:
                confusionSheet.write (intentClasses.index (intent) + 1, intentClasses.index (identifiedIntent) + 1, matrix [intentClasses.index (intent), intentClasses.index (identifiedIntent)])
        rowCnt = 0
        utteranceSheet.write (0, 0, "Utterance")
        utteranceSheet.write (0, 1, "Intent")
        utteranceSheet.write (0, 2, "Identified intent")
        utteranceSheet.write (0, 3, "Confidence")
        utteranceSheet.write (0, 4, "Correct")
        for row in utteranceRows:
            rowCnt += 1
            utteranceSheet.write (rowCnt, 0, row [0])
            utteranceSheet.write (rowCnt, 1, row [1])
            utteranceSheet.write (rowCnt, 2, row [2])
            utteranceSheet.write (rowCnt, 3, row [3])
            utteranceSheet.write (rowCnt, 4, row [4])
            
        outputWorkbook.save (fileName)






