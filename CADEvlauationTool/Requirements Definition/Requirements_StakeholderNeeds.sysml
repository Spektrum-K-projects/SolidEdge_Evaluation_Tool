package 'Stakeholder Needs' {
    private import 'CAD Evaluation Tool Model'::Tool::*;
    private import 'CAD Evaluation Tool Model'::Stakeholder::*;

    package 'Operational Capabilities' {
        requirement def <SN001> AutomatedChecking {
            doc /* The Lead Evaluator needs the tool to automate the checking of Solid Edge examinations.*/
            subject tool : CADEvaluationTool;
            stakeholder evaluator : SH01;
        }
        requirement def <SN002> ScopeLimitation {
            doc /* The Lead Evaluator needs the tool scope to be limited strictly to Part Documents. */
            subject tool : CADEvaluationTool;
            stakeholder evaluator : SH01;
        }
        requirement def <SN003> LanguageInterdependence {
            doc /* The Lead Evaluator needs the tool to function correctly regardless of the Solid Edge language version used by the users (English/German). */
            subject tool : CADEvaluationTool;
            stakeholder evaluator : SH01;
        }
        requirement def <SN004> ToolUI {
            doc /* The Lead Evaluator needs easier access to all tools under a single application window. */
            subject tool : CADEvaluationTool;
            stakeholder evaluator : SH01;
        }
    }

    package 'System Constraint and Design' {
        requirement def <SN005> ToolPerformance {
            doc /* The Developer needs the system to keep the Solid Edge application open throughout the entire execution of the evaluation program. */
            subject tool : CADEvaluationTool;
            stakeholder developer : SH04;
        }
        requirement def <SN006> TextualParsing {
            doc /* The Developer needs the tool to rely on textual data parsing rather than visual parsing. */
            subject tool : CADEvaluationTool;
            stakeholder developer : SH04;
        }
        requirement def <SN007> ModularArchitecture {
            doc /* The Developer needs the extraction tool to be developed with functions structured as separate classes. */
            subject tool : CADEvaluationTool;
            stakeholder developer : SH04;
        }
        requirement def <SN008> ParameterMatching {
            doc /* The Developer needs the code to match features based on parameter values rather than feature names. */
            subject tool : CADEvaluationTool;
            stakeholder developer : SH04;
        }
        requirement def <SN009> MasterReferenceModelID {
            doc /* The Lead Evaluator suggests that a fixed ID (123456) be used as the identifier for the Master Reference Models. */
            subject tool : CADEvaluationTool;
            stakeholder evaluator : SH01;
        }
        requirement def <SN010> FileProcessing {
            doc /* The Lead Evaluator needs a separate module for processing unorganised examination files from the server to a local directory. */
            subject tool : CADEvaluationTool;
            stakeholder evaluator : SH01;
        }
    }

    package 'Data Structuring and Input Handling' {
        requirement def <SN011> StudentModelIdentification {
            doc /* The Lead Evaluator needs student files to be identified using the Matriculation Number and model name. */
            subject tool : CADEvaluationTool;
            stakeholder evaluator : SH01;
        }
        requirement def <SN012> XMLDataFormat {
            doc /* The Developer needs the tool to use XML as the primary format for CAD model data representation. */
            subject tool : CADEvaluationTool;
            stakeholder developer : SH04;
        }
        requirement def <SN013> ConsolidatedOutput {
            doc /* The Developer needs the tool to produce a single XML file per part, containing all extracted information from a single .par model. */
            subject tool : CADEvaluationTool;
            stakeholder developer : SH04;
        }
        requirement def <SN014> DataModelStructuring {
            doc /* The Developer needs the data model to be structured and segregated according to feature. */
            subject tool : CADEvaluationTool;
            stakeholder developer : SH04;
        }
        requirement def <SN015> FeatureDataNesting {
            doc /* The Developer needs extracted data for each feature to be nested inside its respective element. */
            subject tool : CADEvaluationTool;
            stakeholder developer : SH04;
        }
    }

    package 'Core Evaluation and Logic' {
        requirement def <SN016> MasterReference {
            doc /* The Lead Evaluator needs the tool to use the provided Master Reference Model as the reference for grading. */
            subject tool : CADEvaluationTool;
            stakeholder evaluator : SH01;
        }
        requirement def <SN017> SimilarityChecking {
            doc /* The Lead Evaluator needs the tool to check for similarities between student exam models and Master Reference Models. */
            subject tool : CADEvaluationTool;
            stakeholder evaluator : SH01;
        }
        requirement def <SN018> FeatureExtraction {
            doc /* The Lead Evaluator needs the tool to extract data for specific geometric features (ribs, cutouts, extrusions, thin walls). */
            subject tool : CADEvaluationTool;
            stakeholder evaluator : SH01;
        }
        requirement def <SN019> ConstraintChecking {
            doc /* The Lead Evaluator needs the tool to verify whether the model/sketch is fully constrained. */
            subject tool : CADEvaluationTool;
            stakeholder evaluator : SH01;
        }
        requirement def <SN020> PositionCheck {
            doc /* The Lead Evaluator needs to verify the absolute position and location of features relative to the part. */
            subject tool : CADEvaluationTool;
            stakeholder evaluator : SH01;
        }
        requirement def <SN021> SIUnitsVerification {
            doc /* The Lead Evaluator needs the tool to verify that student files use SI units. */
            subject tool : CADEvaluationTool;
            stakeholder evaluator : SH01;
        }
        requirement def <SN022> StyleExtraction {
            doc /* The Lead Evaluator needs the tool to extract colour and style properties from the model. */
            subject tool : CADEvaluationTool;
            stakeholder evaluator : SH01;
        }
        requirement def <SN023> GeneralisedScoring {
            doc /* The Developer needs a generalised evaluation method that is not hard-coded for each examination. */
            subject tool : CADEvaluationTool;
            stakeholder developer : SH04;
        }
    }

    package 'Scoring and Reporting' {
        requirement def <SN024> MaxScoreInput {
            doc /* The Lead Evaluator needs the tool to allow definition of a maximum score for each model. */
            subject tool : CADEvaluationTool;
            stakeholder evaluator : SH01;
        }
        requirement def <SN025> RoundingLogic {
            doc /* The Lead Evaluator needs final scores to be rounded up to the nearest 0.5-point increment. */
            subject tool : CADEvaluationTool;
            stakeholder evaluator : SH01;
        }
        requirement def <SN026> TotalScoreReporting {
            doc /* The Lead Evaluator needs to see individual model scores and a total accumulated score per student. */
            subject tool : CADEvaluationTool;
            stakeholder evaluator : SH01;
        }
        requirement def <SN027> ExcelCompatibility {
            doc /* The Lead Evaluator needs the tool's output to be compatible with Excel and pivot tables. */
            subject tool : CADEvaluationTool;
            stakeholder evaluator : SH01;
        }
    }
}