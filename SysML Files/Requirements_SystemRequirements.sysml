package 'System Requirements' {
	private import RequirementDerivation::*;
    private import 'CAD Evaluation Tool Model'::Tool::*;
    private import 'Stakeholder Needs'::'Operational Capabilities'::*;
    private import 'Stakeholder Needs'::'System Constraint and Design'::*;
    private import 'Stakeholder Needs'::'Data Structuring and Input Handling'::*;
    private import 'Stakeholder Needs'::'Core Evaluation and Logic'::*;
    private import 'Stakeholder Needs'::'Scoring and Reporting'::*;

    requirement def <SR001> PartDocumentScope {
        doc /* The tool shall process only Solid Edge Part Document (.par) files. */
        subject tool : CADEvaluationTool;
    }
    requirement def <SR002> LanguageIndependence {
        doc /* The tool shall produce correct evaluation results regardless of whether the Solid Edge installation is configured in English or German. */
        subject tool : CADEvaluationTool;
    }
    requirement def <SR003> SingleWindowAccess {
        doc /* The tool shall provide access to all evaluation functions through a single application window. */
        subject tool : CADEvaluationTool;
    }
    requirement def <SR004> SolidEdgeSessionPrecondition {
        doc /* The tool shall require an active Solid Edge session to be running prior to initiating evaluation. */
        subject tool : CADEvaluationTool;
    }
    requirement def <SR005> TextualDataParsing {
        doc /* The tool shall extract CAD model data using textual or attribute-based parsing. */
        subject tool : CADEvaluationTool;
    }
    requirement def <SR006> ModularClassStructure {
        doc /* The extraction module shall implement each Solid Edge feature-extraction function as an independently testable class. */
        subject tool : CADEvaluationTool;
    }
    requirement def <SR007> ParameterBasedMatching {
        doc /* The tool shall match corresponding features between a student model and the Master Reference Model based on parameter values rather than feature names. */
        subject tool : CADEvaluationTool;
    }
    requirement def <SR008> MasterReferenceModelID {
        doc /* The tool shall recognise 123456 as the fixed identifier for Master Reference Model files. */
        subject tool : CADEvaluationTool;
    }
    requirement def <SR009> FileTransferModule {
        doc /* The tool shall include a dedicated module for transferring unorganised examination files from the server to a standardised local directory prior to evaluation. */
        subject tool : CADEvaluationTool;
    }
    requirement def <SR010> StudentFileIdentification {
        doc /* The tool shall identify each student file using the student's Matriculation Number and model name. */
        subject tool : CADEvaluationTool;
    }
    requirement def <SR011> XMLDataRepresentation {
        doc /* The tool shall represent extracted CAD model data in XML format. */
        subject tool : CADEvaluationTool;
    }
    requirement def <SR012> ConsolidatedXMLOutput {
        doc /* The tool shall produce a single XML file per part, containing all data extracted from the corresponding .par model. */
        subject tool : CADEvaluationTool;
    }
    requirement def <SR013> FeatureDataSegregation {
        doc /* The tool shall segregate extracted data in the output XML file by feature, such that no two features' data occupy the same structural element. */
        subject tool : CADEvaluationTool;
    }
    requirement def <SR014> FeatureDataNesting {
        doc /* The tool shall nest all data belonging to a single feature within that feature's corresponding parent element in the output XML file. */
        subject tool : CADEvaluationTool;
    }
    requirement def <SR015> MasterReferenceUsage {
        doc /* The tool shall use the Master Reference Model as the reference for evaluating student models. */
        subject tool : CADEvaluationTool;
    }
    requirement def <SR016> RibFeatureExtraction {
        doc /* The tool shall extract feature-specific geometric data for rib features present in the CAD model. */
        subject tool : CADEvaluationTool;
    }
    requirement def <SR017> CutoutFeatureExtraction {
        doc /* The tool shall extract feature-specific geometric data for cutout features present in the CAD model. */
        subject tool : CADEvaluationTool;
    }
    requirement def <SR018> ExtrusionFeatureExtraction {
        doc /* The tool shall extract feature-specific geometric data for extrusion features present in the CAD model. */
        subject tool : CADEvaluationTool;
    }
    requirement def <SR019> ThinWallFeatureExtraction {
        doc /* The tool shall extract feature-specific geometric data for thin wall features present in the CAD model. */
        subject tool : CADEvaluationTool;
    }
    requirement def <SR020> ConstraintVerification {
        doc /* The tool shall verify whether each sketch within the CAD model is fully constrained. */
        subject tool : CADEvaluationTool;
    }
    requirement def <SR021> FeaturePositionVerification {
        doc /* The tool shall verify the position and location of each extracted feature relative to the part. */
        subject tool : CADEvaluationTool;
    }
    requirement def <SR022> SIUnitsVerification {
        doc /* The tool shall verify that the student's CAD model uses SI units. */
        subject tool : CADEvaluationTool;
    }
    requirement def <SR023> StyleExtraction {
        doc /* The tool shall extract colour and style properties from the CAD model. */
        subject tool : CADEvaluationTool;
    }
    requirement def <SR024> GeneralisedEvaluationMethod {
        doc /* The tool shall implement a generalised evaluation method that is not hard-coded to a specific examination. */
        subject tool : CADEvaluationTool;
    }
    requirement def <SR025> FeatureAttributeComparison {
        doc /* The tool shall compare each extracted feature attribute of the student model against the corresponding attribute of the Master Reference Model to determine similarity. */
        subject tool : CADEvaluationTool;
    }
    requirement def <SR026> MaxScoreDefinition {
        doc /* The tool shall allow the maximum score for each model to be defined by the user. */
        subject tool : CADEvaluationTool;
    }
    requirement def <SR027> ScoreRoundingLogic {
        doc /* The tool shall round final scores up to the nearest 0.5 point increment */
        subject tool : CADEvaluationTool;
    }
    requirement def <SR028> ScoreReporting {
        doc /* The tool shall report individual model scores and a total accumulated score for each student. */
        subject tool : CADEvaluationTool;
    }
    requirement def <SR029> ExcelCompatibleOutput {
        doc /* The tool's output shall be compatible with Excel, including pivot table analysis. */
        subject tool : CADEvaluationTool;
    }
    requirement def <SR030> FeatureSequenceExtraction {
        doc /* The tool shall extract feature data in the sequential order in which each feature was created within the CAD model's feature tree. */
        subject tool : CADEvaluationTool;
    }

    #derivation connection def SN002_SR001_Derivation {
        end #original sn002 : SN002;
        end #derive sr001 : SR001;
    }
    #derivation connection def SN003_SR002_Derivation {
        end #original sn003 : SN003;
        end #derive sr002 : SR002;
    }
    #derivation connection def SN004_SR003_Derivation {
        end #original sn004 : SN004;
        end #derive sr003 : SR003;
    }
    #derivation connection def SN005_SR004_Derivation {
        end #original sn005 : SN005;
        end #derive sr004 : SR004;
    }
    #derivation connection def SN006_SR005_Derivation {
        end #original sn006 : SN006;
        end #derive sr005 : SR005;
    }
    #derivation connection def SN007_SR006_Derivation {
        end #original sn007 : SN007;
        end #derive sr006 : SR006;
    }
    #derivation connection def SN008_SR007_Derivation {
        end #original sn008 : SN008;
        end #derive sr007 : SR007;
    }
    #derivation connection def SN009_SR008_Derivation {
        end #original sn009 : SN009;
        end #derive sr008 : SR008;
    }
    #derivation connection def SN010_SR009_Derivation {
        end #original sn010 : SN010;
        end #derive sr009 : SR009;
    }
    #derivation connection def SN011_SR010_Derivation {
        end #original sn011 : SN011;
        end #derive sr010 : SR010;
    }
    #derivation connection def SN012_SR011_Derivation {
        end #original sn012 : SN012;
        end #derive sr011 : SR011;
    }
    #derivation connection def SN013_SR012_Derivation {
        end #original sn013 : SN013;
        end #derive sr012 : SR012;
    }
    #derivation connection def SN014_SR013_Derivation {
        end #original sn014 : SN014;
        end #derive sr013 : SR013;
    }
    #derivation connection def SN015_SR014_Derivation {
        end #original sn015 : SN015;
        end #derive sr014 : SR014;
    }
    #derivation connection def SN016_SR015_Derivation {
        end #original sn016 : SN016;
        end #derive sr015 : SR015;
    }
    #derivation connection def SN018_SR016to019_Derivation {
        end #original sn018 : SN018;
        end #derive sr016 : SR016;
        end #derive sr017 : SR017;
        end #derive sr018 : SR018;
        end #derive sr019 : SR019;
    }
    #derivation connection def SN019_SR020_Derivation {
        end #original sn019 : SN019;
        end #derive sr020 : SR020;
    }
    #derivation connection def SN020_SR021_Derivation {
        end #original sn020 : SN020;
        end #derive sr021 : SR021;
    }
    #derivation connection def SN021_SR022_Derivation {
        end #original sn021 : SN021;
        end #derive sr022 : SR022;
    }
    #derivation connection def SN022_SR023_Derivation {
        end #original sn022 : SN022;
        end #derive sr023 : SR023;
    }
    #derivation connection def SN023_SR024_Derivation {
        end #original sn023 : SN023;
        end #derive sr024 : SR024;
    }
    #derivation connection def SN017_SR025_Derivation {
        end #original sn017 : SN017;
        end #derive sr025 : SR025;
    }
    #derivation connection def SN024_SR026_Derivation {
        end #original sn024 : SN024;
        end #derive sr026 : SR026;
    }
    #derivation connection def SN025_SR027_Derivation {
        end #original sn025 : SN025;
        end #derive sr027 : SR027;
    }
    #derivation connection def SN026_SR028_Derivation {
        end #original sn026 : SN026;
        end #derive sr028 : SR028;
    }
    #derivation connection def SN027_SR029_Derivation {
        end #original sn027 : SN027;
        end #derive sr029 : SR029;
    }
}