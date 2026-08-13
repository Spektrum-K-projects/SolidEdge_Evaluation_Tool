package 'Extraction Implementation' {

    private import 'Architecture Definition'::*;
    private import 'CAD Evaluation Tool Model'::Tool::*;

    item def CADFiles;
    item def XMLDocument;

    part def PartDocumentDataExtractor {
        event occurrence connectSEApplication;
        then event occurrence extractPartDocumentData;
        then event occurrence storeDocumentDataInXML;
        then event occurrence mergePropertiesData;
        then event occurrence mergeFeatureData;
    }
    part def PartPropertiesExtractor {
        event occurrence extractPartProperties;
        then event occurrence storePropertiesInXML;
    }
    part def FeatureDataExtractor {
	    event occurrence extractFeatureData;
	    then event occurrence storeOwnFeatureData;
	    then event occurrence storeFeatureDataInXML;
	}
    part def GeometryDataExtractor {
        event occurrence extractGeometryData;
        then event occurrence storeGeometryDataInXML;
    }
    part def XMLAssembler {
        event occurrence combineXMLData;
        then event occurrence exportXML;
    }

    part def SolidEdgeApplication {
        event occurrence connectExtractionModule;
        then event occurrence runPartFile;
    }

    part def 'Extraction Module' {
        part partDocumentDataExtractor : PartDocumentDataExtractor;
        part partPropertiesDataExtractor : PartPropertiesExtractor;
        part featureDataExtractor : FeatureDataExtractor;
        part geometryDataExtractor : GeometryDataExtractor;
        part xmlAssembler : XMLAssembler;
    }

    part extractionModule : 'Extraction Module';
    part seApplication : SolidEdgeApplication;

    occurrence def 'Document And Properties Extraction' {
        ref part module : 'Extraction Module';
        ref part seApp : SolidEdgeApplication;

        message openCADFile of CADFiles
            from module.partDocumentDataExtractor.connectSEApplication
            to seApp.connectExtractionModule;

        then message runCADFile of CADFiles
            from seApp.runPartFile
            to module.partDocumentDataExtractor.extractPartDocumentData;

        then message documentDataExtract of XMLDocument
            from module.partDocumentDataExtractor.extractPartDocumentData
            to module.partDocumentDataExtractor.storeDocumentDataInXML;

        then message propertiesRequest of CADFiles
            from module.partDocumentDataExtractor.storeDocumentDataInXML
            to module.partPropertiesDataExtractor.extractPartProperties;

        then message propertiesDataExtract of XMLDocument
            from module.partPropertiesDataExtractor.extractPartProperties
            to module.partPropertiesDataExtractor.storePropertiesInXML;

        then message propertiesDataMerge of XMLDocument
            from module.partPropertiesDataExtractor.storePropertiesInXML
            to module.partDocumentDataExtractor.mergePropertiesData;
    }

    occurrence def 'Feature Extraction' {
        ref part module : 'Extraction Module';

        message featureRequest of CADFiles
		    from module.partDocumentDataExtractor.mergePropertiesData
		    to module.featureDataExtractor.extractFeatureData;
		
		then message featureDataExtract of XMLDocument
		    from module.featureDataExtractor.extractFeatureData
		    to module.featureDataExtractor.storeOwnFeatureData;
		
		then message featureGeometryRequest of CADFiles
		    from module.featureDataExtractor.storeOwnFeatureData
		    to module.geometryDataExtractor.extractGeometryData;
		
		then message geometryDataExtract of XMLDocument
		    from module.geometryDataExtractor.extractGeometryData
		    to module.geometryDataExtractor.storeGeometryDataInXML;
		
		then message featureGeometryData of XMLDocument
		    from module.geometryDataExtractor.storeGeometryDataInXML
		    to module.featureDataExtractor.storeFeatureDataInXML;
		
		then message featureDataMerge of XMLDocument
		    from module.featureDataExtractor.storeFeatureDataInXML
		    to module.partDocumentDataExtractor.mergeFeatureData;
		    }

    occurrence def 'XML Data Export' {
        ref part extractionModule : 'Extraction Module';
        ref part seApplication : SolidEdgeApplication;

        occurrence documentPhase : 'Document And Properties Extraction' {
            part redefines module = extractionModule;
            part redefines seApp = seApplication;
        }

        then occurrence featurePhase : 'Feature Extraction'[1..*] ordered {
            part redefines module = extractionModule;
        }

        then message xmlAssembleRequest of CADFiles
            from extractionModule.partDocumentDataExtractor.mergeFeatureData
            to extractionModule.xmlAssembler.combineXMLData;

        then message exportedXMLDocument of XMLDocument
            from extractionModule.xmlAssembler.combineXMLData
            to extractionModule.xmlAssembler.exportXML;
    }
}