package 'CAD Evaluation Tool Model' {

    package Tool {
        part def CADEvaluationTool;
    }

    package Stakeholder {
    	private import Tool::*;
    	
        part def <SH01> 'Lead Evaluator'{
        	doc /* Defines grading criteria, initiates the evaluation process, reviews final scores, and has authority to override results. Creates the Master Reference Model that serves as the baseline for comparison.*/
        }
        part def <SH02> 'Assistant Evaluator'{
        	doc /* Operates the tool to process batches of student files; assists in the grading workflow but does not define criteria.*/
        }
        part def <SH03> 'Model Author'{
        	doc /* Creates the Examination Candidate files that the system must import and analyze. */
        }
        part def <SH04> 'Developer'{
        	doc /* Designs, builds, and validates the tool; defines the initial SE architecture. */
        }
        part def <SH05> 'Contributor'{
        	doc /* Fixes bugs, updates documentation, or develops new modules (e.g., Assembly integration) in the future. */
        }
        
        dependency from SH01 to CADEvaluationTool;
        dependency from SH02 to CADEvaluationTool;
        dependency from SH03 to CADEvaluationTool;
        dependency from SH04 to CADEvaluationTool;
        dependency from SH05 to CADEvaluationTool;
    }            
}