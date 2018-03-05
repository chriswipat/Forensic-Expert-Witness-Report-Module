/*
 * Class ForensicExpertWitnessReport.java of package ForensicExpertwitnessReport
 * 
 * Using this class you are able to retrieve files which have been tagged
 * under a certain name in Autopsy, and add these files and their information
 * into a structured table inside a given Microsoft Word Document.
 * 
 * This class was written for a final year project for
 * the degree of Computer and Digital Forensics BSc (Hons),
 * at Northumbria University in Newcastle, with the aim of
 * aiding in automation, ease and effectivness of digital 
 * forensic practitioners while conducting digital forensic
 * investigations in Autopsy.
 * 
 * @author Chris Wipat
 * @version 19.04.2018
 */

package ForensicExpertWitnessReport;

import javax.swing.JPanel;
import org.sleuthkit.autopsy.report.GeneralReportModule;
import org.sleuthkit.autopsy.report.ReportProgressPanel;
import java.nio.file.Paths;
import org.sleuthkit.autopsy.report.ReportProgressPanel.ReportStatus;
import org.sleuthkit.autopsy.casemodule.services.TagsManager;
import org.sleuthkit.datamodel.ContentTag;
import java.util.ArrayList;
import java.util.logging.Level;
import org.sleuthkit.autopsy.casemodule.Case;
import org.sleuthkit.datamodel.TagName;
import org.sleuthkit.datamodel.TskCoreException;
import java.util.List;

public class ForensicExpertWitnessReport implements GeneralReportModule
{
    // Declare Instance Variables
    public String name = "Forensic Expert Witness Report";
    public String desc = "Add tagged files into a forensic expert witness report.";
    public String filepath = "";
    public String normalizedBaseDir;
    public TagsManager tagsmanager = Case.getCurrentCase().getServices().getTagsManager();
    public ArrayList<ContentTag> TaggedFiles;
    private List<TagName> tagNames;
    private static ForensicExpertWitnessReport instance;   
    private ForensicExpertWitnessReportConfigPanel configPanel;
    
    @Override
    public String getName() {
        return name;
    }

    @Override
    public String getDescription() {
        return desc;
    }

    @Override
    public String getRelativeFilePath() {
        return filepath;
    }
    
    @Override
    public void generateReport(String baseReportDir, ReportProgressPanel progressPanel) {  
        
        // Declare a list of type ContentTag
        List<ContentTag> TaggedFiles = new ArrayList<ContentTag>();
        
        // Retrieve content of tagged file names, set to our declared list
        try {
            TaggedFiles = tagsmanager.getAllContentTags();
        } catch (TskCoreException ex) {
            java.util.logging.Logger.getLogger(ForensicExpertWitnessReport.class.getName()).log(Level.SEVERE, null, ex);
        }
        
        progressPanel.setIndeterminate(false);
        progressPanel.start();
        progressPanel.setMaximumProgress(TaggedFiles.size());
        progressPanel.updateStatusLabel("Reporting files...");
        
        // For each content of tagged file in the ArrayList        
        for (ContentTag CT : TaggedFiles) {
            
//            // skip tags that we are not reporting on 
//            if (passesTagNamesFilter(CT.getName().getDisplayName()) == false) {
//                continue;
//            }
                   
            String fileName;
            try {
                fileName = CT.getContent().getUniquePath();
            } catch (TskCoreException ex) {
                fileName = CT.getContent().getName();
            }
            
            // Report.write(t);
            progressPanel.increment();
        }
  
        // Add the report to the Case, so it is shown in the tree                      
        try {
            normalizedBaseDir = Paths.get(baseReportDir).normalize().toString();
            filepath = Paths.get(normalizedBaseDir, "report.docx").normalize().toString();
            Case.getCurrentCase().addReport(filepath, name, "report.docx");
        } catch (TskCoreException ex) {
            java.util.logging.Logger.getLogger(ForensicExpertWitnessReport.class.getName()).log(Level.SEVERE, null, ex);
        }
        
        progressPanel.complete(ReportStatus.COMPLETE);  

    }

    @Override
    public JPanel getConfigurationPanel() {
        configPanel = new ForensicExpertWitnessReportConfigPanel();
        return configPanel;
    }
    
    // Get the default instance of this report
    public static synchronized ForensicExpertWitnessReport getDefault() {
        if (instance == null) {
            instance = new ForensicExpertWitnessReport();
        }
        return instance;
    }
    
    
    
}