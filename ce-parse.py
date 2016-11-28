#!usr/bin/python

# Author: Heather Young
# Date: Nov 26 2016
# Written for Current Electric Training course parsing
#

import sys
import re
import csv
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

# Argument(s): 
# 	pIndex: an index on the paragraph list
# 	pList: a list of Paragraph Objects
# 	qList: a list containing all parsed questions
# Return: a list of dictionaries, each with question content
def paragraphsToQuestions(pIndex, pList, qDict, qList):
	nextIndex, qDict = parseQuestion(pIndex, pList, qDict)
	# Add dictionary to question list
	print "QDICT IS", qDict
	qList.append(qDict)
	# print "QLIST AFTER APPEND", qList
	# print "passing Index: " + str(nextIndex) + " and recursing"
	# Check to see if we have reached the end of the list
	if nextIndex >= len(pList):
		print str(len(qList))
		print "QLIST AFTER APPEND", qList
		return qList
	else: 
		return paragraphsToQuestions(nextIndex, pList, {}, qList)

# Recursively parses questions from Paragraphs
# Argument(s):
# 	index: index of the text we are expecting to be a question
# 	pList: a list of Paragraph Objects
# 	qDict: a dictionary which will contain the question and corresponding answers
# Returns: a tuple (index of the next question, the filled qDict)
def parseQuestion(index, pList, qDict):
	paragraph = pList[index].text
	print "ATTEMPTING TO PARSE QUESTION\n" + paragraph
	# ignore empty string or Copyright text, also check if the question has at least three words
	if not paragraph or paragraph.isspace() or "Copyright" in paragraph:
		print "CAUGHT BAD PARAGRAPH\n", paragraph
		if index + 1 >= len(pList):
			return index + 1, {}
		else: return parseQuestion(index+1, pList, qDict)
	else:
		qDict['index'] = str(index)
		qDict['question'] = paragraph
		# print "ADDED TO DICT\n", qDict
		return parseAnswers(index + 1, pList, qDict)


# Recursively parse answers for a given question
# Argument(s):
# 	index: index of the text we are expecting to be a answer
# 	pList: list of Paragraph Objects
# 	qDict: a dictionary which will contain the question and corresponding answers
# Returns: a tuple (index of the next paragraph to parse, the filled qDict)
def parseAnswers(index, pList, qDict):
	print "INDEX IS " + str(index)
	paragraph = pList[index].text
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
		if hasattr(result, 'named'): 
			qDict.update(result.named) 
		else: 
			qDict.update(result)
		# print "dict is now ", qDict
		# if the dictionary is complete
		if 'answer4' in qDict:
			print "DICTIONARY COMPLETE"
			nextIndex = index + 1
			return nextIndex, qDict
		# recurse on the rest of the answers
		return parseAnswers(index + 1, pList, qDict)
	else:
		return parseAnswers(index + 1, pList, qDict)

def parseFourAnswersPerLine(answerLine):
	formatString = '{answer1:w}{:s}{:w}.{:s}{answer2:w}{:s}{:w}.{:s}{answer3:w}{:s}{:w}.{:s}{answer4:w}'
	return parse(formatString,answerLine)


def parseTwoAnswersPerLine(index, answerLine):
	firstIdx = index
	secondIdx = index + 2
	pattern = '([ABCD]\.\s+)?(?P<answer' + str(firstIdx) + '>.+?)(\t+)([ABCD]\.\s+)?(?P<answer' + str(secondIdx) + '>.+)(\t+)?'
	result = re.search(pattern, answerLine)
	returnDict = result.groupdict() if hasattr(result, 'groupdict') else None
	return returnDict
	# formatNoEnum = '{}{answer' + str(firstIdx) + ':w}{:s}{:w}.{:s}{answer' + str(secondIndx) + ':w}'
	# formatWithEnum = '{}{answer' + str(firstIdx) + ':w}{:s}{:w}.{:s}{answer' + str(secondIndx) + ':w}'
	# result = parse(formatString, answerLine)

def parseOneAnswerPerLine(index, answerLine):
	formatString = '{answer' + str(index) + '}'
	return parse(formatString, answerLine)

def listQuestions(filename):
	questionList = []
	questionDict = {}
	# get list of Paragraph Objects
	paragraphList = readDoc(filename)
	# truncate the Header content
	return paragraphsToQuestions(firstQuestion(paragraphList), paragraphList, questionDict, questionList)

#write questionList to CSV file
def writeCSV(listQuestions, outputFile):
	with open (outputFile, 'w') as csvfile:
		fieldnames = listQuestions[0].keys()
		writer = csv.DictWriter(csvfile, fieldnames = fieldnames)
		writer.writeheader()
		for q in listQuestions:
			print "\nATTEMPTING WRITE", q.values()
			for fieldname, field in q.iteritems():
				q[fieldname] = q[fieldname].encode('utf-8')
			writer.writerow(q)

def main():
	inputFile = sys.argv[1]
	outputFile = sys.argv[2]
	qList = listQuestions(inputFile)
	return writeCSV(qList, outputFile)

if __name__ == "__main__":
    main()
