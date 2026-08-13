package 'Architecture Definition' {

	private import 'CAD Evaluation Tool Model'::Tool::*;
	private import  'System Requirements'::*;
	
    item def FolderPathData;
    item def MaxScoreRequest;
    item def MaxScoreResponse;

    port def PathInputPort {
        in item sourcePath : FolderPathData;
        in item targetPath : FolderPathData;
    }

    port def XMLInputPort {
        in item studentXMLPath : FolderPathData;
        in item referenceXMLPath : FolderPathData;
    }

    port def ReportOutputPort {
        in item reportOutputPath : FolderPathData;
    }

    port def ScoreInputPopupPort {
        out item request : MaxScoreRequest;
        in item response : MaxScoreResponse;
    }

    action def ExtractFeatures;
    action def ReorganiseFiles;
    action def RequestMaxScores;
    action def CollectMaxScoreInput;
    action def CompareAndScore;

    part def UIShell {
        port toPreprocessing : ~PathInputPort;
        port toExtraction : ~PathInputPort;
        port toComparisonReportingXML : ~XMLInputPort;
        port toComparisonReportingOutput : ~ReportOutputPort;
        port scoreInputPopup : ~ScoreInputPopupPort;
    }

    part def Preprocessing {
        port fromUIShell : PathInputPort;
    }

    part def Extraction {
        port fromUIShell : PathInputPort;
    }

    part def ComparisonReporting {
        port fromUIShellXML : XMLInputPort;
        port fromUIShellOutput : ReportOutputPort;
        port scoreInputPopup : ScoreInputPopupPort;
    }

    part def CADEvaluationTool {

        action performEvaluation {
            action extractReferenceFeatures : ExtractFeatures;
            action reorganiseFiles : ReorganiseFiles [0..1];
            action extractStudentFeatures : ExtractFeatures;
            action requestMaxScores : RequestMaxScores;
            action collectMaxScoreInput : CollectMaxScoreInput;
            action compareAndScore : CompareAndScore;

            first start;
            then fork evaluationStart;

            first evaluationStart then extractReferenceFeatures;
            first evaluationStart then reorganiseFiles;

            first reorganiseFiles then extractStudentFeatures;
            first extractReferenceFeatures then requestMaxScores;
            first requestMaxScores then collectMaxScoreInput;

            first extractStudentFeatures then evaluationReady;
            first collectMaxScoreInput then evaluationReady;
            join evaluationReady;

            first evaluationReady then compareAndScore;
            first compareAndScore then done;
        }

        part uiShell : UIShell {
            perform performEvaluation.collectMaxScoreInput;
        }

        part preprocessing : Preprocessing {
            perform performEvaluation.reorganiseFiles;
        }

        part extraction : Extraction {
            perform performEvaluation.extractReferenceFeatures;
            perform performEvaluation.extractStudentFeatures;
        }

        part comparisonReporting : ComparisonReporting {
            perform performEvaluation.requestMaxScores;
            perform performEvaluation.compareAndScore;
        }

        interface uiShell.toPreprocessing to preprocessing.fromUIShell;
        interface uiShell.toExtraction to extraction.fromUIShell;
        interface uiShell.toComparisonReportingXML to comparisonReporting.fromUIShellXML;
        interface uiShell.toComparisonReportingOutput to comparisonReporting.fromUIShellOutput;
        interface uiShell.scoreInputPopup to comparisonReporting.scoreInputPopup;
        
        requirement sr003 : SR003;
		requirement sr009 : SR009;
        requirement sr010 : SR010;
        requirement sr001 : SR001;
        requirement sr002 : SR002;
        requirement sr004 : SR004;
        requirement sr005 : SR005;
        requirement sr006 : SR006;
        requirement sr008 : SR008;
        requirement sr011 : SR011;
        requirement sr012 : SR012;
        requirement sr013 : SR013;
        requirement sr014 : SR014;
        requirement sr016 : SR016;
        requirement sr017 : SR017;
        requirement sr018 : SR018;
        requirement sr019 : SR019;
        requirement sr023 : SR023;
        requirement sr030 : SR030;
        requirement sr007 : SR007;
        requirement sr015 : SR015;
        requirement sr020 : SR020;
        requirement sr021 : SR021;
        requirement sr022 : SR022;
        requirement sr024 : SR024;
        requirement sr025 : SR025;
        requirement sr026 : SR026;
        requirement sr027 : SR027;
        requirement sr028 : SR028;
        requirement sr029 : SR029;

		allocate sr003 to uiShell;
		
		allocate sr009 to preprocessing;
		allocate sr010 to preprocessing;
		
		allocate sr001 to extraction;
		allocate sr002 to extraction;
		allocate sr004 to extraction;
		allocate sr005 to extraction;
		allocate sr006 to extraction;
		allocate sr008 to extraction;
		allocate sr011 to extraction;
		allocate sr012 to extraction;
		allocate sr013 to extraction;
		allocate sr014 to extraction;
		allocate sr016 to extraction;
		allocate sr017 to extraction;
		allocate sr018 to extraction;
		allocate sr019 to extraction;
		allocate sr023 to extraction;
		allocate sr030 to extraction;
		
		allocate sr007 to comparisonReporting;
		allocate sr015 to comparisonReporting;
		allocate sr020 to comparisonReporting;
		allocate sr021 to comparisonReporting;
		allocate sr022 to comparisonReporting;
		allocate sr024 to comparisonReporting;
		allocate sr025 to comparisonReporting;
		allocate sr026 to comparisonReporting;
		allocate sr027 to comparisonReporting;
		allocate sr028 to comparisonReporting;
		allocate sr029 to comparisonReporting;
    }
}