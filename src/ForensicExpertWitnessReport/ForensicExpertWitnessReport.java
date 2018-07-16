/*
 * Class ForensicExpertWitnessReport.java of package ForensicExpertWitnessReport
 * 
 * Using this class you are able to retrieve files which have been tagged
 * under any given name in Autopsy and add these files and their information
 * into a structured table inside a given Microsoft Word Document, or a pre-existing
 * forensic expert witness report template that comes with this Report Module.
 * 
 * This class was written for a final year project for
 * the degree of Computer and Digital Forensics BSc (Hons),
 * at Northumbria University in Newcastle. This project included
 * the aim of aiding in automation, ease and effectiveness of digital 
 * forensic practitioners whilst conducting digital forensic
 * investigations in Autopsy.
 * 
 * @author Chris Wipat
 * @version 19.04.2018
 */

package ForensicExpertWitnessReport;

import javax.swing.JPanel;
import org.sleuthkit.autopsy.report.GeneralReportModule;
import org.sleuthkit.autopsy.report.ReportProgressPanel;
import org.sleuthkit.autopsy.casemodule.services.TagsManager;
import org.sleuthkit.datamodel.AbstractFile;
import org.sleuthkit.datamodel.Content;
import org.sleuthkit.datamodel.ContentTag;
import java.util.ArrayList;
import java.util.logging.Level;
import org.sleuthkit.autopsy.casemodule.Case;
import org.sleuthkit.datamodel.TagName;
import org.sleuthkit.datamodel.TskCoreException;
import java.util.List;
import javax.swing.JOptionPane;
import org.sleuthkit.autopsy.coreutils.Logger;
import java.nio.file.Paths;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.File;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlCursor;

public class ForensicExpertWitnessReport implements GeneralReportModule {
    
    private final String name = "Forensic Report";
    private final String desc = "Add tagged files into a forensic expert witness report.";
    private String fullpath = "";
    public TagsManager tagsmanager = Case.getCurrentCase().getServices().getTagsManager();
    private List<TagName> tagNames;
    private static ForensicExpertWitnessReport instance;   
    private ForensicExpertWitnessReportConfigPanel configPanel;
    private XWPFDocument ForensicExpertWitnessReport_doc = null;
    private String evidenceHeading = null;    
    private FileOutputStream out = null;
    private String file_extension = "docx";
    private String tableColour = null;

    /**
     * GetName Method
     * First Accessor Method.
     * 
     * Used to return the name of the report module to Autopsy.
     * Necessary while implementing an Autopsy report module.
     * 
     * @return name 
     */
    @Override
    public String getName() {
        return name;
    }    
    
    /**
     * GetDescription Method
     * Second Accessor Method.
     * 
     * Used to return the description of the report module to Autopsy.
     * Necessary while implementing an Autopsy report module.
     * 
     * @return desc 
     */
    @Override
    public String getDescription() {
        return desc;
    }

    /**
     * GetRelativeFilePath Method
     * Third Accessor Method.
     * 
     * Used to return the name of the report which is generated by
     * this report module. Necessary while implementing an Autopsy 
     * report module.
     * 
     * @return "report.docx"
     */
    @Override
    public String getRelativeFilePath() {
        return "report.docx";
    }
    
    /**
     * GenerateReport Method.
     * Main and First mutator method
     * 
     * Uses the selected tag names, target document and evidence header given by 
     * ForensicExpertWitnessReportConfigPanel to retrieve the file information for files
     * under this tag name, and add them into a structured table under the appropriate
     * evidence header inside the given Microsoft Word document.
     * 
     * @param baseReportDir
     * @param progressPanel 
     */
    @Override
    public void generateReport(String baseReportDir, ReportProgressPanel progressPanel) {
        
        // Request inputted configuration details from our GUI panel.
        ForensicExpertWitnessReport_doc = configPanel.getSelectedDocument();
        evidenceHeading = configPanel.getEvidenceHeading();
        file_extension = configPanel.getFileExtension();
        tableColour = configPanel.getTableColour();
                
        // Set the progressPanel to a known amount, start the progressPanel and update it.
        progressPanel.setIndeterminate(false);
        progressPanel.start();
        progressPanel.updateStatusLabel("Adding files...");
        
        // Retrieve the tagsManager from Autopsy
        TagsManager tagsManager = Case.getCurrentCase().getServices().getTagsManager();
        tagNames = configPanel.getSelectedTagNames();
        
        // Create arraylist containing the failed to report tagged files
        ArrayList<String> failedExports = new ArrayList<String>();
         
	// For each tag name in the list of tag names, do the following
        for (TagName tagName : tagNames) {
			
            // Break the loop if the user clicks cancel
            if (progressPanel.getStatus() == ReportProgressPanel.ReportStatus.CANCELED) {
                break;
            }
            
            // Account for false user inputs
            if (ForensicExpertWitnessReport_doc == null) {
                JOptionPane.showMessageDialog(null, "Inputted Document Error.", "Unable to add tagged files to the report", JOptionPane.ERROR_MESSAGE);
                break;
            }
            if (evidenceHeading == null || (evidenceHeading.isEmpty())) {
                JOptionPane.showMessageDialog(null, "Please enter an evidence heading", "Inputted Evidence Heading Error", JOptionPane.ERROR_MESSAGE);
                break;
            }
            if (evidenceHeading.length() < 3) {
                JOptionPane.showMessageDialog(null, "Evidence headings must be 3 characters or longer.", "Inputted Evidence Heading Error", JOptionPane.ERROR_MESSAGE);
                break;
            }
			
            // Try-catch the following, required for retrieving the content of the tagged files.
            try {
                // Request the content of the tagged files by their name and set to a list
                List<ContentTag> tags = tagsManager.getContentTagsByTagName(tagName);

                // Set progress bar to the amount of files we are reporting
                progressPanel.setMaximumProgress(tags.size());
                progressPanel.updateStatusLabel("Adding \"" + tagName.getDisplayName() + "\" files to " + configPanel.getSelectedDocumentName() + "...");

                // Retrieve the paragraphs from the user inputted forensic expert witness report  
                paragraphlist = ForensicExpertWitnessReport_doc.getParagraphs();

                // Convert arraylist to array
                paragraphs = new XWPFParagraph[paragraphlist.size()]; 
                for(int i=0; i<paragraphlist.size(); i++) {
                    paragraphs[i] = paragraphlist.get(i);
                }

                // Declare array of tables to the amount of tagged files retrieved
                tables = new XWPFTable[tags.size()];

                // Count the amount of tables we are creating
                count = 0;

                // For each tagged file, do the following                
                for (ContentTag tag : tags) {

                    // Retrieve the content of the tagged file
                    Content content = tag.getContent();

                    // If the content object relating to this tagged file is an instance of AbstractFile class, do the following.
                    if (content instanceof AbstractFile) {

                        // Update the status label to the current tagged file we are reporting.
                        progressPanel.updateStatusLabel("Adding " + tag.getContent().getName() + " from \"" + tagName.getDisplayName() + "\" to " + configPanel.getSelectedDocumentName() + "...");

                        // Set all variables to blank, to eradicate duplicate entries in table.
                        filename = "";
                        Path = "";
                        md5hash = "";
                        comment = ""; 
                        createdtime = "";
                        modifiedtime = "";
                        accessedtime = "";

                        // Retrieve the File Name, set to variable
                        filename = tag.getContent().getName();                                

                        // Retrieve File Path
                        if (null != ((AbstractFile) content).getLocalAbsPath()) {
                            Path = ((AbstractFile) content).getLocalAbsPath();                                
                        } else {
                            Path = tag.getContent().getUniquePath();                                
                        } 

                        // Retrieve MD5 Hash
                        md5hash = ((AbstractFile) content).getMd5Hash();                                

                        // Retrieve Created Time
                        createdtime = ((AbstractFile) content).getCtimeAsDate();   

                        // Retrieve Modified Time
                        modifiedtime = ((AbstractFile) content).getMtimeAsDate();     

                        // Retrieve Accessed Time
                        accessedtime = ((AbstractFile) content).getAtimeAsDate();  

                        // Retrieve the comment
                        if (tag.getComment() != null) {
                            comment = tag.getComment().trim();
                        }          
                        
                        // Count the amount of evidence headings found
                        heading_count = 0;

                        // Build the Table for this file with the retrieved metadata information
                        buildTables(tags, filename, Path, md5hash, comment, createdtime, modifiedtime, accessedtime);
                        
                        // Display error if the evidence heading was not found & break loop
                        if (heading_count == 0) {
                            JOptionPane.showMessageDialog(null, "Unable to find evidence heading", "Inputted Evidence Heading Error", JOptionPane.ERROR_MESSAGE);
                            break;
                        }

                        // Display error if the evidence heading was not found & break loop
                        if (heading_count > 1 ) {
                            JOptionPane.showMessageDialog(null, "Evidence headings must be unique.", "Multiple entities of headings found", JOptionPane.ERROR_MESSAGE);
                            break;
                        }

                        // Increment the progressPanel every time a file is processed
                        progressPanel.increment();  
                    }
                    // Display an error if the tagged file is not an instance of AbstractFile and thus cannot be written to the report.
                    // This can possibly occur if the tagged file is a directory or if it is unallocated space.
                    else {
                        JOptionPane.showMessageDialog(null, "Unable to add " + tag.getContent().getName() + "to the report.", "Add to Report Error", JOptionPane.ERROR_MESSAGE);
                        failedExports.add(tag.getContent().getName());
                    }
                }

            // Throw exception if we cannot retrieve the content of any of the tagged files
            } catch (TskCoreException ex) {
                Logger.getLogger(ForensicExpertWitnessReport.class.getName()).log(Level.SEVERE, "Error adding files", ex);
                JOptionPane.showMessageDialog(null, "Error getting selected tags for case.", "File Export Error", JOptionPane.ERROR_MESSAGE);
            }
            
            // Write the Document in file system
            try {
                out = new FileOutputStream(new File(baseReportDir + "report." + file_extension));
            } catch(FileNotFoundException e){
                JOptionPane.showMessageDialog(null, "Unable to create new report.", "Create New Report Error", JOptionPane.ERROR_MESSAGE);
                Logger.getLogger(ForensicExpertWitnessReportConfigPanel.class.getName()).log(Level.SEVERE, "Failed to create new report", e);
            }

            // Save the document to disk.                            
            if(out != null) {
                try {
                    ForensicExpertWitnessReport_doc.write(out);
                    out.close();
                } catch(IOException e){
                    JOptionPane.showMessageDialog(null, "Unable to save report.", "Save Report Error", JOptionPane.ERROR_MESSAGE);
                    Logger.getLogger(ForensicExpertWitnessReportConfigPanel.class.getName()).log(Level.SEVERE, "Failed to save report", e);
                }
            }            
        }
        // Manage the failed exports and display to user
        if (!failedExports.isEmpty()) {
            StringBuilder errorMessage = new StringBuilder("Failed to export the following files: ");
            for (int i = 0; i < failedExports.size(); i++) {
                errorMessage.append(failedExports.get(i));
                if (failedExports.size() > 1 && i < failedExports.size() - 1) {
                    errorMessage.append(",");
                }
                if (i == failedExports.size() - 1) {
                    errorMessage.append(".");
                }
            }
            JOptionPane.showMessageDialog(null, errorMessage.toString(), "Hash Export Error", JOptionPane.ERROR_MESSAGE);
        }
        
        // Add the report to the Case, so it is shown in the tree                      
        try {
            fullpath = Paths.get(baseReportDir).normalize().toString();
            Case.getCurrentCase().addReport(fullpath, name, getRelativeFilePath());
        } catch (TskCoreException ex) {
            java.util.logging.Logger.getLogger(ForensicExpertWitnessReport.class.getName()).log(Level.SEVERE, null, ex);
            JOptionPane.showMessageDialog(null, "Unable to add report to report tree", "File Tree Error", JOptionPane.ERROR_MESSAGE);
        }
        
        // Set progress panel status to complete
        progressPanel.complete(ReportProgressPanel.ReportStatus.COMPLETE);
    }
        
    /**
     * Build Tables Method
     * Third mutator method.
     * 
     * Builds table using given information about tagged Autopsy file.
     * 
     * @param tags
     * @param filename
     * @param Path
     * @param md5hash
     * @param comment
     * @param createdtime
     * @param modifiedtime
     * @param accessedtime 
     */
    public void buildTables(List<ContentTag> tags, String filename, String Path, String md5hash, String comment, String createdtime, String modifiedtime, String accessedtime) 
    {
        /**
         * For each paragraph in the forensic expert witness report, do the
         * following. Must not be enhanced loop set to every paragraph in the 
         * document, otherwise java.util.ConcurrentModificationException while 
         * trying to create a new paragraph.
         */
        if(paragraphs != null) {
            for(int i=0; i<paragraphs.length; i++)
            {
                // If the paragraph contains the evidence heading
                if (paragraphs[i].getText() != null && paragraphs[i].getText().contains(evidenceHeading)) {

                    // Count the amount of headings found
                    heading_count++;
                    
                    // If multiple evidence headings are found, break this loop
                    if(heading_count > 1) {
                        break;
                    }

                    // Make sure count is running properly
                    if (count<=tags.size()) {

                        /**
                         * & If this is the first table created, set the cursor to directly after the paragraph object
                         * which contains the evidence heading, and create the table at this point. 
                         * 
                         * Add the table to an array of tables which we have created.
                         */             
                        if(count <1) {
                            cursor = paragraphs[i].getCTP().newCursor();
                            cursor.toNextSibling();
                            table = ForensicExpertWitnessReport_doc.insertNewTbl(cursor);
                            tables[count] = table;
                        }

                        /**
                         * If this is not the first table created, set the cursor directly to after the previous
                         * created table, and create the table at this point.
                         * 
                         * Add the table to an array of tables which we have created
                         */                
                        else {
                            // Set cursor below the previous comment after previous table, if it exists
                            if(para !=  null) {
                                cursor = para.getCTP().newCursor();
                                cursor.toNextSibling();
                                table = ForensicExpertWitnessReport_doc.insertNewTbl(cursor);
                                tables[count] = table;
                            }
                            // If it doesn't exist, set the cursor to 2 siblings after previous table
                            else {
                                cursor = tables[(count-1)].getCTTbl().newCursor();
                                cursor.toNextSibling();
                                cursor.toNextSibling();
                                table = ForensicExpertWitnessReport_doc.insertNewTbl(cursor);
                                tables[count] = table;
                            }                                
                        }                            

                        // Create first row of table // File Name
                        XWPFTableRow tableRowOne = table.getRow(0);
                        tableRowOne.getCell(0).setText("File Name");
                        tableRowOne.getCell(0).setColor(tableColour);
                        tableRowOne.addNewTableCell();
                        if (filename != null) {
                            tableRowOne.getCell(1).setText(filename);
                        }

                        // Create second row of table // File Path
                        XWPFTableRow tableRowTwo = table.createRow();
                        tableRowTwo.getCell(0).setText("File Path");
                        tableRowTwo.getCell(0).setColor(tableColour);
                        if (Path != null) {
                            tableRowTwo.getCell(1).setText(Path);
                        }

                        // Create third row of table // Hash Value
                        XWPFTableRow tableRowThree = table.createRow();
                        tableRowThree.getCell(0).setText("Hash Value");
                        tableRowThree.getCell(0).setColor(tableColour);
                        if (md5hash != null) {
                            tableRowThree.getCell(1).setText(md5hash);
                        }
                        else {
                            tableRowThree.getCell(1).setText("Hashes have not been calculated. Please configure and run an appropriate ingest module.");
                        }

                        // Create fourth row of table // Created time
                        XWPFTableRow tableRowFour = table.createRow();
                        tableRowFour.getCell(0).setText("Created time");
                        tableRowFour.getCell(0).setColor(tableColour);
                        if (Path != null) {
                            tableRowFour.getCell(1).setText(createdtime);
                        }

                        // Create fifth row of table // Modified time
                        XWPFTableRow tableRowFive = table.createRow();
                        tableRowFive.getCell(0).setText("Modified time");
                        tableRowFive.getCell(0).setColor(tableColour);
                        if (Path != null) {
                            tableRowFive.getCell(1).setText(modifiedtime);
                        }

                        // Create sixth row of table // Accessed time
                        XWPFTableRow tableRowSix = table.createRow();
                        tableRowSix.getCell(0).setText("Accessed time");
                        tableRowSix.getCell(0).setColor(tableColour);
                        if (Path != null) {
                            tableRowSix.getCell(1).setText(accessedtime);
                        }

                        // Create paragraph after table // Comment
                        cursor = table.getCTTbl().newCursor();
                        cursor.toNextSibling();
                        if (cursor != null) {
                            para = ForensicExpertWitnessReport_doc.insertNewParagraph(cursor);                                
                        }
                        if ( para != null) {
                            if (comment != null && !(comment.isEmpty())) {                    
                                run = para.createRun();
                                run.setText(comment);
                            }
                            if ((comment == null || comment.isEmpty()) && filename != null) {
                                run = para.createRun();
                                run.setText("This table shows information about \"" +filename + "\"");
                            }  
                        }

                        // Create gap before each table insert
                        cursor = table.getCTTbl().newCursor();
                        if (cursor != null) {
                            para2 = ForensicExpertWitnessReport_doc.insertNewParagraph(cursor);                               
                        }
                        if (para2 != null) {
                            run2 = para2.createRun();
                            run2.setText("");
                        }

                        // Increment the amount of tables created
                        count++;
                    }
                }
            }
        }
    }
    
        /**
     * GetConfigurationPanel Method.
     * Fourth Accessor Method.
     * 
     * First method called by Autopsy to show the GUI of the report module to the user.
     * 
     * @return configPanel
     */
    @Override
    public JPanel getConfigurationPanel() {
        configPanel = new ForensicExpertWitnessReportConfigPanel();
        return configPanel;       
    }
    
    /**
     * GetDefault Method.
     * Fifth Accessor Method.
     * 
     * Get the default instance of this report, used to return an instance of the report
     * back to Autopsy.
     * 
     * @return instance
     */
    public static synchronized ForensicExpertWitnessReport getDefault() {
        if (instance == null) {
            instance = new ForensicExpertWitnessReport();
        }
        return instance;
    }
    
    // Further Variable Declaration //GEN-BEGIN:variables
    private String filename;
    private String Path;
    private String md5hash;
    private String comment;
    private String createdtime;
    private String modifiedtime;
    private String accessedtime;
    private int count;
    private int heading_count;
    private List<XWPFParagraph> paragraphlist;
    private XWPFTable[] tables;
    private XWPFParagraph[] paragraphs;
    private XmlCursor cursor;
    private XWPFTable table;
    private XWPFRun run;
    private XWPFRun run2;
    private XWPFParagraph para;
    private XWPFParagraph para2;
    // End of variables declaration//GEN-END:variables 
}