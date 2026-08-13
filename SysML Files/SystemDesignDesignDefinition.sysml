package 'Design Definition' {
	
	private import 'Architecture Definition'::*;
	private import ScalarValues::*;

    item def FeatureXMLNode;
    item def SimilarityScore;
		
	port def FeatureNodePort {
	    out item featureNode : FeatureXMLNode;
	}

    port def SimilarityScorePort {
        out item score : SimilarityScore;
    }

    action def BuildFeatureNode;
	action def AssembleDocument;
    action def MatchAndCompareFeatures;
    action def ComputeAndReportScore;
    
    calc def ComputeScore {
        in similarityRatio : Real;
        in maxScore : Real;
        return result : Real = similarityRatio * maxScore;
    }
    
    action buildNode : BuildFeatureNode;
    action assemble : AssembleDocument;
	
	part def SolidEdgeExtractor {
	    port nodeOutput : FeatureNodePort;
	    perform buildNode;
	}
	
	part def XMLAssembler {
	    port nodeInput : ~FeatureNodePort;
	    perform assemble;
	}
	
	part def ExtractionDesign :> Extraction {
	    part solidEdgeExtractor : SolidEdgeExtractor;
	    part xmlAssembler : XMLAssembler;
	
	    interface featureNodeHandoff connect solidEdgeExtractor.nodeOutput to xmlAssembler.nodeInput;
	
	    succession modelData first solidEdgeExtractor.buildNode then xmlAssembler.assemble;
	}
	
	action match : MatchAndCompareFeatures;
    action report : ComputeAndReportScore;

    part def Comparator {
        port scoreOutput : SimilarityScorePort;
        perform match;
    }

    part def ScoringReporter {
        port scoreInput : ~SimilarityScorePort;
        perform report;
    }

    part def ComparisonReportingDesign :> ComparisonReporting {
        part comparator : Comparator;
        part scoringReporter : ScoringReporter;

        interface scoreGenerator connect comparator.scoreOutput to scoringReporter.scoreInput;

        succession similarityScore first comparator.match then scoringReporter.report;
    }

    part cadEvaluationToolDesign : CADEvaluationTool {
        part  :>> extraction : ExtractionDesign;
        part  :>> comparisonReporting : ComparisonReportingDesign;
    }
}