package 'Verification Definition' {
    private import VerificationCases::*;
    private import VerificationMethodKind::*;
    private import ScalarValues::String;
    private import 'Design Definition'::*;

    item def VerifiedFeatureNode :> FeatureXMLNode {
        attribute reportedStatus : String;
    }

    requirement featureStatusCheck {
        doc /* Refines SR006: for each extracted feature, SolidEdgeExtractor shall report
              status using the Solid Edge FeatureStatusConstants enumeration rather than
              an unconverted raw integer value. */

        subject extractor : SolidEdgeExtractor;
        in reportedStatus : String;
        in expectedStatus : String;

        require constraint {
            reportedStatus == expectedStatus
        }
    }

    verification def FeatureStatusVerification {
        subject testExtractor : SolidEdgeExtractor;

        objective featureStatusVerificationObjective {
            verify featureStatusCheck;
        }

        action collectData {
            in part testExtractor : SolidEdgeExtractor = FeatureStatusVerification::testExtractor;
            out reportedStatus : String;
        }

        action evaluateData {
            in reportedStatus : String = collectData.reportedStatus;
            out verdict : VerdictKind = PassIf(
                featureStatusCheck(
                    extractor = testExtractor,
                    reportedStatus = reportedStatus,
                    expectedStatus = "seFeatureOK"
                )
            );
        }

        @VerificationMethod { kind = (test, analyze); }

        return verdict : VerdictKind = evaluateData.verdict;
    }

    part verificationContext {
        part extractorUnderTest : SolidEdgeExtractor;

        verification featureStatusTest : FeatureStatusVerification {
            subject testExtractor = extractorUnderTest;
        }
    }
}