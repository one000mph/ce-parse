#!usr/bin/python

# Author: Heather Young
# Date: Nov 26 2016
# Written for Current Electric Training course parsing
# Reads a Microsoft Word document, parses questions and 
# corresponding answers and then exports the questions
# to a .csv file
#

import sys
import re
import csv
import time
from parse import *
from docx import *

# Reads a Microsoft Word document using the docx Python library
# Argument(s): 
#	filepath: the file to be read
# Return: a list of Paragraph Objects
def readDoc(filepath):
	document = Document(filepath)
	paragraphList = document.paragraphs
	return paragraphList

# removes the header paragraph from the course
# Argument(s):
# 	pList: a list of Paragraph Objects
# Return: the index of the first question
def firstQuestion(pList):
	for idx, paragraph in enumerate(pList):
		strParagraph = paragraph.text
		# We assume that the header format remains consistent
		if "Expire" in strParagraph:
			return idx + 1


# removes the headers from the answer sheet
# Argument(s):
# 	pList: a list of Paragraph Objects
# Return: the index of the first answer ['A', 'B', 'C', 'D']
def firstAnswer(pList):
	for idx, paragraph in enumerate(pList):
		# strip numerical list, underscores, and whitespace
		strParagraph = (paragraph.text).strip('1234567890.\t_ ')
		# print "\nparagraph", strParagraph
		if strParagraph:
			strParagraph = strParagraph[0]
			if len(strParagraph) < 2:
				result = re.match('[ABCD]', strParagraph)
				if result != None and result.group(0):
					print "results", result.group(0)
					return idx


# Argument(s): 
# 	qDict: a dictionary containing a question and its answers
# 	aLetter: the letter of the correct answer
# Return: the qDict with 'answer4' removed and the other answers shifted up by one
def reOrderAnswers(qDict, answerLetter):
	answerIndices = {'A':0, 'B':1, 'C':2, 'D':3}
	answerFields = ['answer1', 'answer2', 'answer3', 'answer4']
	answerIndex = answerIndices[answerLetter]
	qDict['correct-answer'] = qDict[answerFields[answerIndex]]
	for i in range(answerIndex, 3):
		qDict[answerFields[i]] = qDict[answerFields[i+1]]
	del qDict['answer4']
	return qDict


# Argument(s): 
# 	qDict: a dictionary containing a question and its answers
# 	aIndex: an index on the answer list
#   aIndex: a list of answers paragraphs
# Return: the qDict with the correct answer under the field 'correct-answer'
def selectAnswerAndReference(qDict, aIndex, aList):
	sanitizedData = aList[aIndex].text.lstrip('1234567890.\t_ ')
	answerLetter = sanitizedData[0]
	reference = sanitizedData[1:].strip('\t_ ')
	# print "LETTER: ", answerLetter
	# print "REFERENCE: ", reference
	qDict['reference'] = reference
	return reOrderAnswers(qDict, answerLetter)

# Argument(s): 
# 	pIndex: an index on the paragraph list
# 	pList: a list of Paragraph Objects
# 	qDict: a dictionary containing a question and its answers
# 	qList: a list containing all parsed questions
#   aIndex: an index on the answer list
# Return: a list of dictionaries, each with question content
def paragraphsToQuestions(pIndex, pList, qDict, qList, aIndex, aList):
	nextIndex, qDict = parseQuestion(pIndex, pList, qDict)
	# Add dictionary to question list
	if qDict:
		qDict = selectAnswerAndReference(qDict, aIndex, aList)
	qList.append(qDict)
	# Check to see if we have reached the end of the list
	if nextIndex >= len(pList):
		return qList
	else: 
		return paragraphsToQuestions(nextIndex, pList, {}, qList, aIndex+1, aList)

# Recursively parses questions from Paragraphs
# Argument(s):
# 	index: index of the text we are expecting to be a question
# 	pList: a list of Paragraph Objects
# 	qDict: a dictionary which will contain the question and corresponding answers
# Returns: a tuple (index of the next question, the filled qDict)
def parseQuestion(index, pList, qDict):
	paragraph = pList[index].text
	if paragraph and not paragraph.isspace() and len(paragraph.split(' ')) < 4: # suspect an invalid question
		print "POSSIBLE INVALID QUESTION: " + paragraph
		exit(0)
	# ignore empty string or Copyright text, also check if the question has at least three words
	if not paragraph or paragraph.isspace() or "Copyright" in paragraph:
		# print "CAUGHT BAD PARAGRAPH\n", paragraph
		if index + 1 >= len(pList):
			return index + 1, {}
		else: return parseQuestion(index+1, pList, qDict)
	else:
		qDict['index'] = str(index)
		qDict['question'] = paragraph
		return parseAnswers(index + 1, pList, qDict)


# Recursively parse answers for a given question
# Argument(s):
# 	index: index of the text we are expecting to be a answer
# 	pList: list of Paragraph Objects
# 	qDict: a dictionary which will contain the question and corresponding answers
# Returns: a tuple (index of the next paragraph to parse, the filled qDict)
def parseAnswers(index, pList, qDict):
	paragraph = pList[index].text.strip(' \t')
	if len(paragraph.strip()) is not 0: # discard empty strings
		# print "answer might be " + paragraph
		result = parseFourAnswersPerLine(paragraph)
		if not result:
			# try parsing two answers to a line
			aIdx = 1 if 'answer1' not in qDict else 2 # assign relative answer index
			# print "AIDX IS " + str(aIdx)
			result = parseTwoAnswersPerLine(aIdx, paragraph)
			if not result:
				# try parsing one answer per line
				aIdx = 4 if 'answer3' in qDict else (3 if 'answer2' in qDict else (2 if 'answer1' in qDict else 1))
				result = parseOneAnswerPerLine(aIdx, paragraph)
				if not result: # Handle error case
					print "ANSWER PARSE FAILED ON #" + str(index)
					exit(0)
		# add any parsed answers to qDict
		qDict.update(result) 
		# if the dictionary is complete
		if 'answer4' in qDict:
			nextIndex = index + 1
			return nextIndex, qDict
		# recurse on the rest of the answers
		return parseAnswers(index + 1, pList, qDict)
	else:
		return parseAnswers(index + 1, pList, qDict)


# Argument(s):
# 	answerLine: the string we are parsing
# Returns: the parsed answers in a dictionary with keys {'answer1': ..., 'answer2':...}
def parseFourAnswersPerLine(answerLine):
	formatString = '{answer1:w}{:s}{:w}.{:s}{answer2:w}{:s}{:w}.{:s}{answer3:w}{:s}{:w}.{:s}{answer4:w}'
	parsedResult = parse(formatString,answerLine)
	result = parsedResult if hasattr(parsedResult, 'named') else None
	return result

# Argument(s):
# 	index: the index of the answer [1: parse answers A and C, 2: parse answers B and D]
# 	answerLine: the string we are parsing
# Returns: the parsed answers in a dictionary with keys {'answer1': ..., 'answer2':...}
def parseTwoAnswersPerLine(index, answerLine):
	firstIdx = index
	secondIdx = index + 2
	pattern = '([ABCD]\.\s+)?(?P<answer' + str(firstIdx) + '>.+?)(\t+)([ABCD]\.\s+)?(?P<answer' + str(secondIdx) + '>.+)(\t+)?'
	result = re.search(pattern, answerLine)
	returnDict = result.groupdict() if hasattr(result, 'groupdict') else None
	return returnDict

# Argument(s):
# 	index: the index of the answer [1: A, 2: B, 3: C, 4: D]
# 	answerLine: the string we are parsing
# Returns: the parsed answers in a dictionary with keys {'answer1': ..., }
def parseOneAnswerPerLine(index, answerLine):
	pattern = '([ABCD]\.\s+)?(?P<answer' + str(index) + '>.+)'
	result = re.search(pattern, answerLine)
	returnDict = result.groupdict() if hasattr(result, 'groupdict') else None
	return returnDict


# Write questionList to CSV file
# Argument(s):
# 	listQuestions: a list of question dictionaries
# 	outputFile: a .csv file to write to
# Returns: Nothing
def writeCSV(listQuestions, outputFile):
	with open (outputFile, 'w') as csvfile:
		# set order of columns
		fieldnames = ['index', 'question', 'reference', 'correct-answer', 'answer1', 'answer2', 'answer3']
		writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
		writer.writeheader()
		for q in listQuestions:
			for fieldname, field in q.iteritems():
				q[fieldname] = q[fieldname].encode('utf-8')
			writer.writerow(q)
	return None

def main():
	inputFile = sys.argv[1]
	answerFile = sys.argv[2]
	outputFile = sys.argv[3]
	questionList = [] # initialize list to contain question dictionaries
	questionDict = {} # initialize first question dictionary
	paragraphList = readDoc(inputFile) # get list of Paragraph Objects
	answerList = readDoc(answerFile) # list of answers
	firstQIndex = firstQuestion(paragraphList) # get index of first real question, ignore header
	firstAIndex = firstAnswer(answerList)
	print firstAIndex
	qList = paragraphsToQuestions(firstQIndex, paragraphList, questionDict, questionList, firstAIndex, answerList) # fill questionList
	return writeCSV(qList, outputFile)

if __name__ == "__main__":
    main()
