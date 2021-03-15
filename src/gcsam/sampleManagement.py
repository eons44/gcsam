
# Sample Management is the heart of these scripts. This file defines what FAMEs, Samples, Lines, etc. are and what can be done with them.
# The only reason this file should be changed is if there are necessary features that must be added.


#A TrendIdentifier is a base class for any datum type that one would like to find trends between.
#The members of this class are helpful labels along with the ability to invalidate a datum.
class TrendIdentifier:
    m_number = 0
    m_name = "INVALID NAME"
    m_colorID = "" #a unique color to help identify trends
    m_uniqueID = 0
    m_valid = True #this makes it easy to create a new object without having to call MakeValid
    def __init__(self):
        self.m_number = 0
        self.m_name = "INVALID NAME"
        self.m_colorID = ""
        self.m_uniqueID = 0
        self.m_valid = True

    #RETURNS: m_valid
    def IsValid(self):
        return self.m_valid == True

    #Sets m_valid to true
    def MakeValid(self):
        self.m_valid = True

    #Sets m_valid to false.
    def Invalidate(self):
        self.m_valid = False

#A FAME object contains many helpful ways of analyzing the outputs of a gas chromatograph.
#These primarily include the retention time and peak area.
#However, additional members exist that can be calculated later.
#   For example, the percentFA requires using the peak area along with the mass of the sample and standard used.
#FAMEs also contain a name that can be selected from a set of labeled retention times.
#   The means of identifying unknown FAMEs is rather complex, but the method FAME1IsBetterMatchThan(...) (defined outside of this class) is helpful in executing ApplyLabelsWithClosestMatchFrom(...)
#For more details on FAMEs manipulation, see FAMEsContainer
class FAME (TrendIdentifier):
    #NOTE: the FAME's m_number will be set to the m_uniqueID of its parent when used within a LineManager
    m_retentionTime = 0
    m_peakArea = 0
    m_percentFA = 0
    m_percentOfTotalFA = 0
    m_ugFA = 0
    m_nameMatchDiscrepancy = 0 #how close the RT matches that of a labeled FAME
    m_bestMatch = False
    def __init__(self):
        super(TrendIdentifier,self).__init__()
        self.m_retentionTime = 0
        self.m_peakArea = 0
        self.m_percentFA = 0
        self.m_percentOfTotalFA = 0
        self.m_ugFA = 0
        self.m_nameMatchDiscrepancy = 0
        self.m_bestMatch = False

    #Use a formula to pick out which FAMEs are the best fits for their labels.
    #RETURNS: a boolean favoring self (true) or compare (false)
    def IsBetterMatchThan(self, compare):
        selfValue = 0
        compareValue = 0
        if(self.m_peakArea>0 and compare.m_peakArea>0):
            selfValue = self.m_peakArea
            compareValue = compare.m_peakArea
        #WHY IS THIS HERE?
        # elif(self.m_percentFA>0 and compare.m_percentFA>0):
        #     selfValue = self.m_percentFA
        #     compareValue = compare.m_percentFA
        else:
            print("Can't find a suitable comparison between", self.m_name, "and", compare.m_name)
            return False;
        if(self.m_nameMatchDiscrepancy != 0):
            selfValue /= abs(self.m_nameMatchDiscrepancy)
        if(compare.m_nameMatchDiscrepancy != 0):
            compareValue /= abs(compare.m_nameMatchDiscrepancy)
        # print(self.m_name, "(", self.m_retentionTime, ",", self.m_peakArea, ")", 
        #     ":", selfValue, ",", compare.m_name, "(", compare.m_retentionTime, 
        #     ",", compare.m_peakArea, ")", ":", compareValue)
        return selfValue > compareValue

#FAMEsContainer is a base class for all objects that contain FAMEs
#This object has a lot of helpful methods for manipulating FAMEs.
#For details on each of these methods, please see the function definition below.
class FAMEsContainer:
    m_mgFADW = 0 #mg dry weight of tissue used for FAMEs
    m_mgStd = 0 #mg of standard used in the sample
    m_totalWeightFraction = 0 #mg FAMEs to mg tissue
    m_fames = []
    def __init__(self):
        self.m_mgFADW = 0
        self.m_mgStd = 0
        self.m_totalWeightFraction = 0
        self.m_totalFA = 0
        self.m_fames = []

    #Adds a FAME to *this
    def AddFAME(self,fame):
        self.m_fames.append(fame)

    #RETURNS: a FAME of the given name, an invalid FAME if none found.
    def GetFAME(self,name):
        for f in self.m_fames:
            if(f.m_name == name and f.m_bestMatch):
                return f
        ret = FAME()
        ret.Invalidate()
        return ret

    #Calculates the percentFA for each fame based on comparison to the peak area of the std.
    #REQUIREMENTS:
    #   1. all fames be labeled
    #   2. the stdName provided matches a FAME within *this
    #   3. m_mgFADW and m_mgStd are valid numbers.
    def CalculatePercentFA(self,stdName):
        std = self.GetFAME(stdName)
        if(not std.IsValid()):
            print(self.m_name, "- PercentFA not calculated:", stdName, "could not be found.")
            return
        if(self.m_mgStd == 0):
            print("Invalid mg standard in",self.m_name)
            return
        if(self.m_mgFADW == 0):
            print("Invalid tissue weight in",self.m_name)
            return

        self.m_totalWeightFraction = 0
        weightFraction = self.m_mgStd/self.m_mgFADW
        # print(self.m_name, "weight fraction is", weightFraction)
        for f in self.m_fames:
            if (not f.IsValid()):
                continue
            if (f.m_name == "INVALID NAME" or not f.m_bestMatch):
                continue
            if(f.m_name == std.m_name):
                continue
            f.m_percentFA = f.m_peakArea/std.m_peakArea * weightFraction
            # print(f.m_name, "has PA", f.m_peakArea, std.m_name, "has PA", std.m_peakArea, "%FA:", "{0:.10f}".format(f.m_percentFA))
            self.m_totalWeightFraction += f.m_peakArea
            # print(self.m_name, "using", f.m_name, "for total fraction (", self.m_totalWeightFraction, ")")
        self.m_totalWeightFraction /= std.m_peakArea 
        # print(self.m_name, "total weight fraction is", self.m_totalWeightFraction)
        self.m_totalWeightFraction *= weightFraction
        # print(self.m_name, "total weight fraction is", self.m_totalWeightFraction)

    #Removes all FAMEs in toRem from *this.
    #RETURNS: the FAMEs removed
    def RemoveFAMES(self,toRem):
        self.m_fames = [f for f in self.m_fames if f not in toRem]
        return toRem

    #Adds all Samples from otherFAMEsContainer to *this.
    #If there are duplicate FAMES they are removed here.
    #RETURNS: the FAMEs removed
    def ImportDataFrom(self, otherFAMEsContainer):
        self.m_fames.extend(otherFAMEsContainer.m_fames);
        toRem = []
        alreadyProcessed = []
        for f1 in self.m_fames:
            skip = False
            for fp in alreadyProcessed:
                if(f1.m_retentionTime == fp):
                    skip = True
                    break
            if(skip):
                continue
            for f2 in self.m_fames:
                if(f1 is not f2 and f1.m_retentionTime == f2.m_retentionTime):
                    print("Removing duplicate FAME with retention time", f2.m_retentionTime)
                    toRem.append(f2)
                    alreadyProcessed.append(f1.m_retentionTime)
        return self.RemoveFAMES(toRem)

    #IMPORTANT: for the following methods to work, FAMEs must have a valid percentFA value.

    #Removes all FAMEs from m_fames that have a percentFA content less/more than percentFA
    #RETURNS: the FAMEs removed
    def RemoveFAMEsLowerThanPercentFA(self, percentFA):
        toRem = []
        for f in self.m_fames:
            if(f.m_percentFA < percentFA):
                toRem.append(f)
        return self.RemoveFAMES(toRem)
    def RemoveFAMEsHigherThanPercentFA(self, percentFA):
        toRem = []
        for f in self.m_fames:
            if(f.m_percentFA > percentFA):
                toRem.append(f)
        return self.RemoveFAMES(toRem)

    #Removes all FAMEs from m_fames that have a percentFA% of less/more than percent
    #This is confusing because the total %FA in tissue is <1, so this refers to the %%FA, which should total to 1.
    #RETURNS: the FAMEs removed
    def RemoveFAMEsLowerThanPercent(self, percent):
        percent /= 100 #convert percent to fraction
        t = self.GetTotalPercentFA()
        toRem = []
        for f in self.m_fames:
            if(f.m_percentFA/t < percent):
                toRem.append(f)
        return self.RemoveFAMES(toRem)
    def RemoveFAMEsHigherThanPercent(self, percent):
        percent /= 100 #convert percent to fraction
        t = self.GetTotalPercentFA()
        toRem = []
        for f in self.m_fames:
            if(f.m_percentFA/t > percent):
                toRem.append(f)
        return self.RemoveFAMES(toRem)

    #RETURNS: the sum of m_percentFA for all m_fames
    def GetTotalPercentFA(self, bestMatch = False, standardName = ""):
        t = 0
        for f in self.m_fames:
            if(not bestMatch or (f.m_bestMatch and f.m_name != standardName)):
                t += f.m_percentFA
        return t

    #RETURNS: the lowest/highest m_percentFA from m_fames
    def GetLowestPercentFA(self):
        lPercentFA = 0
        for f in self.m_fames:
            if(f.m_percentFA < lPercentFA):
                lPercentFA = f.m_percentFA
        return lPercentFA
    def GetHighestPercentFA(self):
        hPercentFA = 0
        for f in self.m_fames:
            if(f.m_percentFA > hPercentFA):
                hPercentFA = f.m_percentFA
        return hPercentFA

    #Sets the m_percentOfTotalFA of each FAME to m_percentFA/GetTotalPercentFA(...)
    def CalculatePercentOfTotalFA(self, standardName):
        totalFA = self.GetTotalPercentFA(True,standardName)
        if(totalFA ==0):
            print("Total FA in",self.m_name,"is 0")
            return

        for f in self.m_fames:
            f.m_percentOfTotalFA = f.m_percentFA/totalFA

    #RETURNS: the lowest/highest retention time of FAMEs in *this
    #based on how the data is read in, this should be m_fames[0]
    def GetLowestRT(self):
        if(len(self.m_fames)):
            return self.m_fames[0].m_retentionTime
        return 0
    def GetHighestRT(self):
        if(len(self.m_fames)):
            return self.m_fames[len(self.m_fames)-1].m_retentionTime
        return 0

    #RETURNS: the smallest gap between retention times in m_fames
    def GetSmallestRTGap(self):
        sm = 1000000 #too big
        for i in range(len(self.m_fames)):
            if(i == len(self.m_fames)-1):
                break #we look at i and i+1, so Remove last i
            dRT = abs(self.m_fames[i].m_retentionTime - self.m_fames[i+1].m_retentionTime)
            if(dRT < sm):
                sm = dRT
        return sm

    #RETURNS: an empty, invalid FAME.
    def InvalidFAME(self):
        ret = FAME()
        ret.Invalidate()
        return ret

    #RETURNS: the next RT beyond starting RT that has a percentFA that is higher than that of both RTs next to it.
    #RETURNS: InvalidFame() if the requested value does not exist
    #startingRT will be adjusted to the first valid RT that is >= to startingRT
    def GetNextLocalMaxRT(self,startingRT):
        #check corner cases first
        if(startingRT>= self.GetHighestRT()):
            return self.InvalidFAME()

        for i in range(len(self.m_fames)):
            if(self.m_fames[i].m_retentionTime>startingRT):
                if(i ==0):
                    if(self.m_fames[i].m_percentFA > self.m_fames[i+1].m_percentFA):
                        return self.m_fames[i]
                    else:
                        continue
                if(i ==len(self.m_fames)-1):
                    if(self.m_fames[i].m_percentFA < self.m_fames[i-1].m_percentFA):
                        return self.InvalidFAME()
                    else:
                        return self.m_fames[len(self.m_fames)-1]
                if(self.m_fames[i].m_percentFA > self.m_fames[i+1].m_percentFA and self.m_fames[i].m_percentFA > self.m_fames[i-1].m_percentFA):
                    return self.m_fames[i]

    #RETURNS: an array of all FAMEs in this that are given by GetNextLocalMaxRT
    def GetAllLocalMaxFAMEs(self):
        startingRT = self.GetLowestRT()
        ret = FAMEsContainer()
        while(True):
            tempF = self.GetNextLocalMaxRT(startingRT)
            if(tempF.IsValid()):
                ret.AddFAME(tempF)
                startingRT = tempF.m_retentionTime
            else:
                return ret
            #print("Found local max ( RT",tempF.m_retentionTime,'FA',tempF.m_percentFA,")")

    #Removes all FAMEs in *this that are not returned by GetAllLocalMaxFAMEs()
    def TrimFAMEsToLocalMaximaOnly(self):
        self.m_fames = self.GetAllLocalMaxFAMEs().m_fames

    #Changes m_name to that of the FAME to be that of the labeldFame with the closest retention time.
    #IMPORTANT: This requires that all labeled FAMEs be in RT order!
    def ApplyLabelsWithClosestMatchFrom(self, labeledFames):
        acceptableGapLow = 1 #BIG but not too big.. #FIXME: calculate this?
        for i in range(len(labeledFames.m_fames)):
            if(i!= len(labeledFames.m_fames)-1):
                acceptableGapHigh = abs(labeledFames.m_fames[i+1].m_retentionTime - labeledFames.m_fames[i].m_retentionTime)/2
                # print("Next acceptable gap is", acceptableGapHigh, "making range",labeledFames.m_fames[i].m_retentionTime - acceptableGapLow,"to",labeledFames.m_fames[i].m_retentionTime + acceptableGapHigh)
            else:
                acceptableGapHigh = 0
            for f in self.m_fames:
                if(
                    f.m_retentionTime > labeledFames.m_fames[i].m_retentionTime 
                    - acceptableGapLow 
                    and 
                    f.m_retentionTime <= labeledFames.m_fames[i].m_retentionTime 
                    + acceptableGapHigh
                    ):
                    f.m_name = labeledFames.m_fames[i].m_name
                    f.m_nameMatchDiscrepancy = f.m_retentionTime - labeledFames.m_fames[i].m_retentionTime
            acceptableGapLow = acceptableGapHigh

    #Once FAMEs have labels, this method may be called to find which one "best" matches its label.
    #This is useful iff 2+ FAMEs have the same label, but are not the same.
    #This method relies on IsBetterMatchThan(...) to determine what a "best" match is.
    def FindBestMatchingFAMEs(self):
        currentName = ""
        matchToBeat = FAME()
        matchToBeat.Invalidate()
        for f in self.m_fames:
            #find local min for m_nameMatchDiscrepancy
            if(f.m_name != currentName):
                currentName = f.m_name
                matchToBeat = f
                f.m_bestMatch = True
            elif( not matchToBeat.IsValid() or f.IsBetterMatchThan(matchToBeat)):
                matchToBeat.m_bestMatch = False
                matchToBeat = f;
                f.m_bestMatch = True

    #Removes all FAMEs with the name "INVALID NAME"
    #RETURNS: the removed FAMEs
    def RemoveAllUnlabeledFAMEs(self):
        toRem = []
        for f in self.m_fames:
            if(f.m_name =="INVALID NAME"):
                toRem.append(f)
        return self.RemoveFAMES(toRem)

#ExpressionLevel is a simple class that measures expression level, primarily on western blots.
#This class has no methods and only a coarse rank or more precise "value".
class ExpressionLevel:
    m_westernRank = 0
    m_westernValue = 0
    def __init__(self):
        self.m_westernRank = 0
        self.m_westernValue = 0

#A Sample is a TrendIdentifier, a FAMEsContainer, and has an ExpressionLevel.
class Sample (TrendIdentifier, FAMEsContainer):
    m_expressionLevel = ExpressionLevel()
    m_zygosity = ""
    def __init__(self):
        super(TrendIdentifier,self).__init__()
        super(FAMEsContainer,self).__init__()
        self.m_expressionLevel = ExpressionLevel()
        self.m_zygosity = ""

#A SamplesContainer is a base class for all objects that contain samples.
#Currently, SamplesContainers only support access methods to a set of samples.
class SamplesContainer:
    m_samples = []
    def __init__(self):
        self.m_samples = []

    #Adds a sample to *this
    def AddSample(self,sample):
        sample.m_number = len(self.m_samples)
        self.m_samples.append(sample)

    #RETURNS: the sample from *this with the given name or an invalid sample if none found.
    def GetSampleByName(self, sampleName):
        for s in self.m_samples:
            if(s.m_name ==sampleName):
                return s
        ret = Sample()
        ret.Invalidate()
        return ret

    #RETURNS: the sample from *this with the given number or an invalid sample if none found.
    def GetSampleByNumber(self, sampleNumber):
        for s in self.m_samples:
            if(s.m_number ==sampleNumber):
                return s
        ret = Sample()
        ret.Invalidate()
        return ret

    #Removes all Samples in toRem from *this.
    #RETURNS: the Samples removed
    def RemoveSamples(self,toRem):
        self.m_samples = [s for s in self.m_samples if s not in toRem]
        return toRem

    #Adds all Samples from otherSampleContainer to *this.
    #If there are duplicate FAMES, etc. they are combined here.
    #RETURNS: excess samples removed
    def ImportDataFrom(self, otherSampleContainer):
        self.m_samples.extend(otherSampleContainer.m_samples);
        toRem = []
        alreadyProcessed = []
        for p1 in self.m_samples:
            skip = False
            for pp in alreadyProcessed:
                if(p1.m_name == pp):
                    skip = True
                    break
            if(skip):
                continue
            for p2 in self.m_samples:
                if(p1 is not p2 and p1.m_name == p2.m_name):
                    print("Adding data from and removing duplicate sample", p2.m_name)
                    p1.ImportDataFrom(p2)
                    toRem.append(p2)
                    alreadyProcessed.append(p1.m_name)
        return self.RemoveSamples(toRem)

#A line is a SamplesContainer aswell as a TrendIdentifier and has no other special methods.
class Line (TrendIdentifier, SamplesContainer):
    def __init__(self):
        super(SamplesContainer,self).__init__()
        super(TrendIdentifier,self).__init__()

#A LineManager contains a set of Lines and is the primary means of interacting with all members within Line objects.
#For example, if one wanted to change the name of a FAME, they would get the FAME by first asking the LineManager for the appropriate sample, then query the FAME and change its name.
#A LineManager also assigns a uniqueID to all TrendIdentifiers within it.
class LineManager:
    m_nextID = 0
    m_lines = []
    def __init__(self):
        self.m_nextID = 0
        self.m_lines = []

    #RETURNS: a new unique ID
    def GetNextID(self):
        ret = self.m_nextID
        self.m_nextID+= 1
        return ret

    #RETURNS: a line of the given name, an invalid line if none found.
    def GetLine(self, lineName):
        for l in self.m_lines:
            if(l.m_name ==lineName):
                return l
        ret = Line()
        ret.Invalidate()
        return ret

    #Adds a Line to *this
    def AddLine(self, line):
        self.m_lines.append(line)

    #Removes all Lines in toRem from *this.
    #RETURNS: the Lines removed
    def RemoveLines(self,toRem):
        self.m_lines = [l for l in self.m_lines if l not in toRem]
        return toRem

    def RemoveAllTheseLines(self, namesToRemove):
        toRem = []
        for l in self.m_lines:
            for name in namesToRemove:
                if(l.m_name == name):
                    toRem.append(l)
                    break
        self.RemoveLines(toRem)

    def RemoveAllLinesExcept(self, namesToKeep):
        toKeep = []
        for l in self.m_lines:
            for name in namesToKeep:
                if(l.m_name == name):
                    toKeep.append(l)
                    break
        self.m_lines = [l for l in self.m_lines if l in toKeep]

    #Sets the m_uniqueIDs of all TrendIdentifiers within *this.
    #This method should be called once all TrendIdentifiers are loaded in *this.
    def SetUniqueIDs(self):
        for l in self.m_lines:
            l.m_uniqueID = self.GetNextID()
            for s in l.m_samples:
                s.m_uniqueID = self.GetNextID()
                for f in s.m_fames:
                    f.m_uniqueID = self.GetNextID()
                    f.m_number = s.m_uniqueID

    #Searches through all lines to find a sample of the given name.
    #RETURNS: the sample from a Line in *this with the given name or an invalid sample if none found.
    def GetSampleByName(self, sampleName):
        for l in self.m_lines:
            ret = l.GetSampleByName(sampleName)
            if(ret.IsValid()):
                return ret
        ret = Sample()
        ret.Invalidate()
        return ret

    #Adds all Lines from otherLineManager to *this.
    #If there are duplicate lines, samples, etc. they are combined here.
    #RETURNS: the Lines removed.
    def ImportDataFrom(self, otherLineManager):
        self.m_lines.extend(otherLineManager.m_lines)
        # print(self.m_lines)
        # return
        toRem = []
        alreadyProcessed = []
        for l1 in self.m_lines:
            skip = False
            for lp in alreadyProcessed:
                if(l1.m_name == lp):
                    skip = True
                    break
            if(skip):
                continue
            for l2 in self.m_lines:
                if(l1 is not l2 and l1.m_name == l2.m_name):
                    print("Adding data from and removing duplicate line", l2.m_name)
                    l1.ImportDataFrom(l2)
                    toRem.append(l2)
                    alreadyProcessed.append(l1.m_name)
        return self.RemoveLines(toRem)

    #Prints some information on the data in *this.
    #Currently, these data only include counts of each Line, Sample, and FAME.
    def Analyze(self):
        sampleCount = 0;
        famesCount = 0;
        unlabeledFamesCount = 0
        for l in self.m_lines:
            #print("There are",len(l.m_samples),"samples in",l.m_name)
            sampleCount += len(l.m_samples)
            for s in l.m_samples:
                #print("There are",len(s.m_fames),"FAMEs in sample", s.m_number, "(",s.m_mgFADW, "mg DW )")
                famesCount += len(s.m_fames)
        print("There are",len(self.m_lines),"lines")
        print("There are",sampleCount,"samples")
        print("There are",famesCount,"fames")

    #Labels all FAMEs within all Samples within all Lines within *this according to the given set of known FAMEs.
    def LabelFAMEsWith(self,labeledFAMEs):
        for l in self.m_lines:
            for s in l.m_samples:
                s.ApplyLabelsWithClosestMatchFrom(labeledFAMEs)
                s.FindBestMatchingFAMEs()

    #Calculates the percent FA for all FAMEs within all Samples within all Lines within *this by comparison with the given standardName.
    #REQUIREMENTS:
    #   1. all fames be labeled
    #   2. the stdName provided matches a FAME within each Sample
    #   3. m_mgFADW and m_mgStd for each Sample are valid numbers.
    def CalculatePercentFAForAllSamples(self, standardName):
        for l in self.m_lines:
            for s in l.m_samples:
                s.CalculatePercentFA(standardName)

    #Calculates the percent of total PercentFA for all FAMEs within all Samples within all Lines within *this by comparison with the given standardName.
    #Requires that each FAME have a valid m_percentFA
    def CalculatePercentOfTotalFAForAllSamples(self, standardName):
        for l in self.m_lines:
            for s in l.m_samples:
                s.CalculatePercentOfTotalFA(standardName)

    #Removes any FAMEs within *this with a m_percentFA/totalPercentFA lower than lowestPercent or higher than highestPercent.
    #same for value, where value is in percentFA
    def TrimFAMEsTo(self, lowestPercent, highestPercent, lowestValue, highestValue):
        remT = 0
        for l in self.m_lines:
            remL = 0
            for s in l.m_samples:
                remP = 0
                remP += len(s.RemoveFAMEsLowerThanPercent(lowestPercent))
                remP += len(s.RemoveFAMEsHigherThanPercent(highestPercent))
                remP += len(s.RemoveFAMEsLowerThanPercentFA(lowestValue))
                remP += len(s.RemoveFAMEsHigherThanPercentFA(highestValue))
                remL += remP
            remT += remL
        print("Removed",remT,"FAMEs")

    #Removes all FAMEs within *this that are not local maxima.
    def TrimFAMEsToLocalMaxima(self):
        for l in self.m_lines:
            for s in l.m_samples:
                s.TrimFAMEsToLocalMaximaOnly()

    #Removes all FAMEs within *this that have the name "INVALID NAME"
    def RemoveUnlabeledFAMEs(self):
        for l in self.m_lines:
            for s in l.m_samples:
                s.RemoveAllUnlabeledFAMEs()

    #RETURNS: all samples in *this as an array of SampleLinePairs
    def GetAllSamples(self):
        allSamples = []
        for l in self.m_lines:
            for s in l.m_samples:
                newPair = SampleLinePair()
                newPair.m_sample = s
                newPair.m_line = l
                allSamples.append(newPair)
        return allSamples

#A FAMESamplePair is a simple class containing a FAME and a Sample.
class FAMESamplePair:
    m_fame = FAME()
    m_sample = Sample()
    def __init__(self):
        self.m_fame = FAME()
        self.m_sample = Sample()

#A SampleLinePair is a simple class containing a Sample and a Line.
class SampleLinePair:
    m_sample = Sample()
    m_line = Line()
    def __init__(self):
        self.m_sample = Sample()
        self.m_line = Line()
