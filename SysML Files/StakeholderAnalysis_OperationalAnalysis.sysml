package 'Operational Analysis' {
	private import 'CAD Evaluation Tool Model'::Tool::*;
	private import 'CAD Evaluation Tool Model'::Stakeholder::*;
	
    package 'Tool Operation'{
    
        use case def 'Define Grading Criteria' {
            subject tool : CADEvaluationTool;
            actor evaluator : SH01;
        }

        use case def 'Initiate Evaluation' {
            subject tool : CADEvaluationTool;
            actor evaluator : SH01;
        }

        use case def 'Review Final Scores' {
            subject tool : CADEvaluationTool;
            actor evaluator : SH01;
        }

        use case def 'Override Results' {
            subject tool : CADEvaluationTool;
            actor evaluator : SH01;
        }
        
        use case def 'Reference Model Creation' {
        	subject tool : CADEvaluationTool;
            actor evaluator : SH01;
        }

        use case def 'Process Student Files' {
            subject tool : CADEvaluationTool;
            actor assistant : SH02;
        }
        
        use case def 'Grading Assistance' {
            subject tool : CADEvaluationTool;
            actor assistant : SH02;
        }

        use case def 'Submit Examination Model' {
            subject tool : CADEvaluationTool;
            actor author : SH03;
        }
    }

    package 'Tool Development'{
        use case def 'Define System Requirements' {
            subject tool : CADEvaluationTool;
            actor developer : SH04;
        }
        
        use case def 'Define Verification Criteria' {
            subject tool : CADEvaluationTool;
            actor developer : SH04;
        }

        use case def 'Extend Requirements Baseline' {
            subject tool : CADEvaluationTool;
            actor contributor : SH05;
        }

        use case def 'Propose Future Requirements' {
            subject tool : CADEvaluationTool;
            actor contributor : SH05;
        }
    }
}