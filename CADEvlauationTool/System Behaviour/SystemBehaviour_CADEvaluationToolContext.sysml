package 'CAD Tool Environment' {

    private import 'CAD Evaluation Tool Model'::Tool::*;
    private import 'CAD Evaluation Tool Model'::Stakeholder::*;

    part def 'Solid Edge Application' {
        doc /* The external CAD application the tool depends on. */
    }
    part def 'Examination File Server' {
        doc /* The server from which unorganised examination files are retrieved. */
    }

    package 'System Context' {
        part evaluator : SH01;
        part tool : CADEvaluationTool;
        part seApp : 'Solid Edge Application';
        part fileServer : 'Examination File Server';

        dependency from evaluator to tool {
            doc /* Exchanges Reference Model File and Evaluation Report. */
        }
        dependency from fileServer to tool {
            doc /* Supplies Student Model File. */
        }
        dependency from tool to seApp {
            doc /* Exchanges API Commands and Geometry Data. */
        }
    }
}