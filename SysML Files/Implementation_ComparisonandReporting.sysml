package 'Comparison and Reporting Implementation' {

    item def PartIdentifier;
    item def MatchedReference;
    item def ComparisonData;
    item def SimilarityRatio;
    item def FinalScore;

    part def FileMatcher {
    event occurrence receiveMatchRequest;
    then event occurrence tryExactMatch;
    then event occurrence fallBackFuzzyMatch;
    then event occurrence returnMatchedReference;
	}

    part def Comparator {
        event occurrence compareElements;
        then event occurrence recursiveUnorderedMatch;
        then event occurrence returnComparisonResults;
    }

    part def ScoreCalculation {
        event occurrence calculateSimilarity;
        then event occurrence clampToMaxPoints;
        then event occurrence roundUpToNearestHalf;
        then event occurrence returnFinalScore;
    }

    part def ReportGen {
        event occurrence generateReport;
        then event occurrence writeExcelReport;
    }

    part def Controller {
        event occurrence requestMatch;
        then event occurrence receiveMatch;
        then event occurrence requestComparison;
        then event occurrence receiveComparison;
        then event occurrence requestScore;
        then event occurrence receiveScore;
        then event occurrence requestReport;
    }

    part controller : Controller;
    part fileMatcher : FileMatcher;
    part comparator : Comparator;
    part scoreCalculation : ScoreCalculation;
    part reportGen : ReportGen;

    occurrence def 'Comparison and Scoring' {
        ref part controller : Controller;
        ref part fileMatcher : FileMatcher;
        ref part comparator : Comparator;
        ref part scoreCalculation : ScoreCalculation;
        ref part reportGen : ReportGen;

        message matchRequest of PartIdentifier
		    from controller.requestMatch
		    to fileMatcher.receiveMatchRequest;
		
		then message attemptExactMatch of PartIdentifier
		    from fileMatcher.receiveMatchRequest
		    to fileMatcher.tryExactMatch;
		
		then message fuzzyMatchFallback of PartIdentifier
		    from fileMatcher.tryExactMatch
		    to fileMatcher.fallBackFuzzyMatch;
		
		then message matchedReference of MatchedReference
		    from fileMatcher.returnMatchedReference
		    to controller.receiveMatch;

        then message compareRequest of MatchedReference
            from controller.requestComparison
            to comparator.compareElements;

        then message recursiveCompare of ComparisonData
            from comparator.compareElements
            to comparator.recursiveUnorderedMatch;

        then message comparisonResults of ComparisonData
            from comparator.returnComparisonResults
            to controller.receiveComparison;

        then message scoreRequest of ComparisonData
            from controller.requestScore
            to scoreCalculation.calculateSimilarity;

        then message clampScore of SimilarityRatio
            from scoreCalculation.calculateSimilarity
            to scoreCalculation.clampToMaxPoints;

        then message ceilingScore of SimilarityRatio
            from scoreCalculation.clampToMaxPoints
            to scoreCalculation.roundUpToNearestHalf;

        then message finalScore of FinalScore
            from scoreCalculation.returnFinalScore
            to controller.receiveScore;

        then message reportRequest of FinalScore
            from controller.requestReport
            to reportGen.generateReport;

        then message writeExcelReport of FinalScore
            from reportGen.generateReport
            to reportGen.writeExcelReport;
    }

}