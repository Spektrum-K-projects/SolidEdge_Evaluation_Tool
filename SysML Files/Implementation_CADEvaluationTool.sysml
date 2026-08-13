package 'CAD Evaluation Tool Implementation' {

    private import 'Architecture Definition'::*;

    item def BatchComparisonRequestData;

    part def UIShellSequence :> UIShell {
        event occurrence requestPreprocessing;
        then event occurrence receiveDirectoryReady;
        then event occurrence requestReferenceExtraction;
        then event occurrence receiveReferenceReady;
        then event occurrence requestBatchExtraction;
        then event occurrence receiveStudentXMLReady;
        then event occurrence requestPartsList;
        then event occurrence receivePartsList;
        then event occurrence beginCollectScores;
        then event occurrence finishCollectScores;
        then event occurrence requestComparison;
        then event occurrence receiveExcelReady;
    }

    part def PreprocessingSequence :> Preprocessing {
        event occurrence reorganiseFiles;
        then event occurrence directoryReady;
    }

    part def ExtractionSequence :> Extraction {
        event occurrence extractReferenceModel;
        then event occurrence referenceXMLReady;
        then event occurrence extractStudentBatch;
        then event occurrence studentXMLReady;
    }

    part def ComparisonReportingSequence :> ComparisonReporting {
        event occurrence listParts;
        then event occurrence returnPartList;
        then event occurrence compareAndScore;
        then event occurrence writeExcelReport;
    }

    part uiShell : UIShellSequence;
    part preprocessing : PreprocessingSequence;
    part extraction : ExtractionSequence;
    part comparisonReporting : ComparisonReportingSequence;

    occurrence def 'CAD Evaluation Tool Operation' {
        ref part uiShell : UIShellSequence;
        ref part preprocessing : PreprocessingSequence;
        ref part extraction : ExtractionSequence;
        ref part comparisonReporting : ComparisonReportingSequence;
        
        message extractReferenceRequest of FolderPathData
            from uiShell.requestReferenceExtraction
            to extraction.extractReferenceModel;

        then message referenceReady of FolderPathData
            from extraction.referenceXMLReady
            to uiShell.receiveReferenceReady;

        then message preprocessRequest of FolderPathData
            from uiShell.requestPreprocessing
            to preprocessing.reorganiseFiles;

        then message directoryReady of FolderPathData
            from preprocessing.directoryReady
            to uiShell.receiveDirectoryReady;

        then message batchExtractionRequest of FolderPathData
            from uiShell.requestBatchExtraction
            to extraction.extractStudentBatch;

        then message studentXMLReady of FolderPathData
            from extraction.studentXMLReady
            to uiShell.receiveStudentXMLReady;

        then message partsListRequest of MaxScoreRequest
            from uiShell.requestPartsList
            to comparisonReporting.listParts;

        then message partsList of MaxScoreResponse
            from comparisonReporting.returnPartList
            to uiShell.receivePartsList;

        then message collectScores of MaxScoreResponse
            from uiShell.beginCollectScores
            to uiShell.finishCollectScores;

        then message comparisonRequest of BatchComparisonRequestData
            from uiShell.requestComparison
            to comparisonReporting.compareAndScore;

        then message excelReportReady of FolderPathData
            from comparisonReporting.writeExcelReport
            to uiShell.receiveExcelReady;
    }

}