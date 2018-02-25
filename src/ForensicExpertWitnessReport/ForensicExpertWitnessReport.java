package ForensicExpertWitnessReport;

/*
 * Class ForensicExpertWitnessReport.java of package ForensicExpertwitnessReport
 * 
 * Using this class you are able to ...
 * 
 * @author Chris Wipat
 * @version 21.02.2018
 */

import javax.swing.JPanel;
import org.sleuthkit.autopsy.report.GeneralReportModule;
import org.sleuthkit.autopsy.report.ReportProgressPanel;
import java.util.ArrayList;
import java.util.List;
import com.aspose.words.Document;

import java.lang.System;
import java.nio.file.InvalidPathException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.logging.Level;
import org.sleuthkit.datamodel.TskData;
import org.sleuthkit.autopsy.casemodule.Case;
import org.sleuthkit.autopsy.coreutils.Logger;
import org.sleuthkit.autopsy.report.GeneralReportModuleAdapter;
import org.sleuthkit.autopsy.report.ReportProgressPanel.ReportStatus;
import org.sleuthkit.autopsy.casemodule.services.FileManager;
import org.sleuthkit.datamodel.TskCoreException;
import org.sleuthkit.autopsy.casemodule.services.TagsManager;
import org.sleuthkit.datamodel.ContentTag;
import org.sleuthkit.datamodel.TagName;

import org.sleuthkit.autopsy.report.ReportGenerator.TableReportsWorker;


/*
* First Accessor method 
* 
*/
public class ForensicExpertWitnessReport implements GeneralReportModule
{
    // Declare Instance Variables
    public String name = "ForensicExpertWitnessReport";
    public String desc = "The most awesome module!";
    public String filepath;
    public String normalizedBaseDir;
    public TagsManager tagsmanager;
    public ArrayList<ContentTag> TaggedFiles;
    
    // Static instance of this report
    private static ForensicExpertWitnessReport instance;   

    @Override
    public void generateReport(String baseReportDir, ReportProgressPanel pnl) 
    {       
        // Declare a list of type ContentTag
        List<ContentTag> TaggedFiles = new ArrayList<ContentTag>();
        
        // Retrieve content of tagged files, set to our declared list
        try {
        TaggedFiles = tagsmanager.getAllContentTags();
        } catch (TskCoreException ex) {
            java.util.logging.Logger.getLogger(ForensicExpertWitnessReport.class.getName()).log(Level.SEVERE, null, ex);
        }
        
        pnl.setIndeterminate(false);
        pnl.start();
        pnl.setMaximumProgress(TaggedFiles.size());
        
        // For each content of tagged file in the ArrayList        
        for (ContentTag CT : TaggedFiles) {
            
            // skip tags that we are not reporting on 
            if (passesTagNamesFilter(CT.getName().getDisplayName()) == false) {
                continue;
            }
                   
            String fileName;
            try {
                fileName = CT.getContent().getUniquePath();
            } catch (TskCoreException ex) {
                fileName = CT.getContent().getName();
            }
            
            // Report.write(t);
            pnl.increment();
        }
  
        // Add the report to the Case, so it is shown in the tree                      
        try {
            normalizedBaseDir = Paths.get(baseReportDir).normalize().toString();
            filepath = Paths.get(normalizedBaseDir, "report.docx").normalize().toString();
            Case.getCurrentCase().addReport(filepath, name, "report.docx");
        } catch (TskCoreException ex) {
            java.util.logging.Logger.getLogger(ForensicExpertWitnessReport.class.getName()).log(Level.SEVERE, null, ex);
        }
        
        pnl.complete(ReportStatus.COMPLETE);
        
        
        
    }
    
    private boolean passesTagNamesFilter(String tagName) {
        return tagNamesFilter.isEmpty() || tagNamesFilter.contains(tagName);
    }

    @Override
    public JPanel getConfigurationPanel() {
        // Return null for now
        return null;
    }

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
        return normalizedBaseDir;
    }
    
    // Get the default instance of this report
    public static synchronized ForensicExpertWitnessReport getDefault() {
        if (instance == null) {
            instance = new ForensicExpertWitnessReport();
        }
        return instance;
    }
    
    
    
}