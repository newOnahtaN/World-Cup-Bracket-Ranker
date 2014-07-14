from xlrd import open_workbook
from xlwt import Workbook

#Nathan Owen, ncowen@email.wm.edu


#Needs for use: Two excel sheets, both in very particular formates.  One MUST be named World Cup Rankings and the other
#MUST be named World Cup Scores.  See Kelvin Abrokwa for details about excel sheet format.


#-----------------------------------------------------------------------------------------------------------------------


#Algorthim Outline

# Where each bracket is an order on teams, 1-32, 1 being the best and 32 being the worst,
# for each football match, the difference between the favorable* team's score and the unfavored team's score will be 
# multiplied against the distance in between the two teams found on the bracket.  For instance, one might rank Team A
# as number 7 and team B as number 23.  Result of 1-3 (A-B) would yield a multiplier of -2, and because the distance
# between the teams is 16, this game yields a subtraction of 32 points from this brackets overall accuracy score.
# If a match is tied, then the multiplier will automatically be -0.5 times the distance between the teams. 

# *A team is favorable in a match if they are higher in a bracket then the competing team. Converse for unfavorable.



#-----------------------------------------------------------------------------------------------------------------------
#Misc functions

def findMultiplier(score, firstMinusSecond):
	first = float(score[0])
	second = float(score[2])

	if firstMinusSecond:
		return first - second

	else:
		return second - first



#-----------------------------------------------------------------------------------------------------------------------
#Collect brackets and matches data

brackets = open_workbook('World Cup Rankings.xlsx')
matches = open_workbook('World Cup Scores.xlsx')


#Brackets will be of type python list, element 0 is bracket author name, element 1 is bracket name,
#and 2-33 are the teams in order from best (2) to worst (33)

listOfBrackets = []
for sheet in brackets.sheets():

	for row in range(sheet.nrows):
		bracketList = []

		for column in range(sheet.ncols):
			bracketList += [sheet.cell(row,column).value]

		if 'Your Name' not in bracketList:
			#print bracketList[0]
			listOfBrackets += [bracketList]


#print '\n\n\n'
#print listOfBrackets
#print '\n\n\n'


# Matches will be of type python list where the first two elements are the teams, the third the winner, and the fourth
# the score. 


listOfMatches = []
for sheet in matches.sheets():

	for row in range(sheet.nrows):
		matchList = []

		for column in range(sheet.ncols):
			matchList += [sheet.cell(row,column).value]

		if 'tbd' not in matchList:
			#print matchList
			listOfMatches += [matchList]


#print listOfMatches

#-----------------------------------------------------------------------------------------------------------------------
#Run Algorithm

textFile = open('Bracket Results.txt', 'w')
scoreList = []
bracketDict = {}

for bracket in listOfBrackets:

	numUnrecordedMatches = 0
	bracketScore = 0

	for match in listOfMatches:
		#For each match, add or subtract from the person's score. 

		teamOne = match[0]
		teamTwo = match[1]
		score = match[3]

		#print '\n'
		#print teamOne
		#print teamTwo
		#print score
		

		
		try:
			#Determine whether the first or second team is 'prefered', and then subtract to get multiplier accordingly.
			teamOneRank = bracket.index(teamOne)
			teamTwoRank = bracket.index(teamTwo)

			if teamOneRank < teamTwoRank:
				firstMinusSecond = True
			else:
				firstMinusSecond = False



			distance = float(abs(teamOneRank - teamTwoRank))

			if distance == 0:
				multiplier = -0.5
			else:
				multiplier = findMultiplier(score, firstMinusSecond)

			#print teamOneRank
			#print teamTwoRank
			#print distance
			#print multiplier
			#print distance * multiplier

			bracketScore += (float(distance) * float(multiplier))

		except:
			#print "Bracket " + bracket[0] + " is incomplete because it does not contain or correctly record either " + \
					#teamOne + " or " + teamTwo

			#textFile.write("Bracket " + bracket[0] + " is incomplete because it does not contain or correctly record either " + \
					#teamOne + " or " + teamTwo + '\n')

			numUnrecordedMatches += 1

	bracket[2] = numUnrecordedMatches
	
	if bracketScore not in scoreList:
		bracketDict[bracketScore] = bracket
		scoreList += [bracketScore]
	else:
		bracketScore += .1

		if bracketScore not in scoreList:

			bracketDict[bracketScore] = bracket
			scoreList += [bracketScore]

		else:

			bracketScore += .1
			bracketDict[bracketScore] = bracket
			scoreList += [bracketScore]

scoreList.sort(reverse = True)
#print scoreList


#-----------------------------------------------------------------------------------------------------------------------
#Show results


rank = 1
for score in scoreList:
	bracket = bracketDict[score]

	print 'Rank: {:<2} Name: {:<25s} Bracket Name: {:<40s} Score: {:<15} # of Matches not recorded due to misentry: {:<2}'.\
	format(rank, bracket[0], bracket[1], score, bracket[2])

	textFile.write('Rank: {:<2} Name: {:<25s} Bracket Name: {:<40s} Score: {:<15} # of Matches not recorded due to misentry: {:<2}'.\
	format(rank, bracket[0], bracket[1], score, bracket[2]))
	textFile.write('\n')

	rank += 1

#end = input("Press any key to end this program.")
textFile.close()






		