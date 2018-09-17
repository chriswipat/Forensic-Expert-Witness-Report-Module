/*
 * Forensic Expert Witness Report
 * 
 * This module adds tagged evidence directly into structured and styled tables 
 * inside a forensic expert witness report, allowing the selection of forensic 
 * expert witness reports or coming with three pre-existing forensic expert witness 
 * report templates to choose from. 

 * Contact: Chris Wipat [chris <at> cwipat [dot] co [dot] uk]
 *
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * See LICENSE or <http://www.gnu.org/licenses/>.
 *
 * ###################################
 *
 * Class ForensicReport.java of package ForensicExpertWitnessReport
 * 
 * Using this class you are able to retrieve files which have been tagged under any 
 * given name in Autopsy and add the information about these files into structured 
 * and coloured tables inside any given Microsoft Word Document. The colouring of 
 * these tables and the Microsoft Word Document to report into is given by the 
 * configuration panel. The configuration panel allows the selection of forensic
 * expert witness reports and allows the selection of three forensic expert witness 
 * report templates to choose from.
 * 
 * This Report Module was written for a final year project for the degree
 * of Computer and Digital Forensics BSc (Hons) at Northumbria University 
 * in Newcastle and improved for OSDFCon, 2018. The final year project at
 * Northumbria University included the aim of aiding in the automation, 
 * ease, and effectiveness of digital forensic practitioners whilst 
 * conducting digital forensic investigations in Autopsy.
 *
 * https://github.com/chriswipat/forensic_expert_witness_report_module
 * 
 * @author Chris Wipat
 * @version 17.09.2018
 */

package ForensicReport;

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
import java.io.FileInputStream;
import java.math.BigInteger;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlCursor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSpacing;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblWidth;
import org.sleuthkit.autopsy.datamodel.ContentUtils;

public class ForensicReport implements GeneralReportModule {
    
    private final String name = "Forensic Report";
    private final String desc = "Add tagged files into a forensic expert witness report.";
    private String fullpath = "";
    public TagsManager tagsmanager = Case.getCurrentCase().getServices().getTagsManager();
    private List<TagName> tagNames;
    private static ForensicReport instance;   
    private ForensicReportConfigPanel configPanel;
    private XWPFDocument ForensicReport_doc = null;
    private String evidenceHeading = null;    
    private FileOutputStream out = null;
    private String file_extension = "docx";
    private String tableColour = null;
    private FileInputStream image_is = null;
    private final String fontColour = "ffffff";
//    private XWPFDocument Example_Table;    

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
     * ForensicReportConfigPanel to retrieve the file information for files
     * under this tag name, and add them into a structured table under the appropriate
     * evidence header inside the given Microsoft Word document.
     * 
     * @param baseReportDir
     * @param progressPanel 
     */
    @Override
    public void generateReport(String baseReportDir, ReportProgressPanel progressPanel) {
        
        // Retrieve inputted configuration details from our GUI panel.
        ForensicReport_doc = configPanel.getSelectedDocument();
        evidenceHeading = configPanel.getEvidenceHeading();
        file_extension = configPanel.getFileExtension();
        tableColour = configPanel.getTableColour();
//        Example_Table = configPanel.getExampleTable();
                
        // Set the progressPanel to a known amount, start the progressPanel and update it.
        progressPanel.setIndeterminate(false);
        progressPanel.start();
        progressPanel.updateStatusLabel("Adding files...");
        
        // Retrieve the tagsManager from Autopsy
        TagsManager tagsManager = Case.getCurrentCase().getServices().getTagsManager();
        tagNames = configPanel.getSelectedTagNames();
        
        // Create arraylist containing the failed to report tagged files
        ArrayList<String> failedExports = new ArrayList<String>();
        
        // Create the list containing the type of files which we want to extract the content of into the report
        List<String> img_exts = new ArrayList<String>();
        img_exts.add("jpg"); img_exts.add("gif"); img_exts.add("jpeg"); img_exts.add("png");
              
	// For each tag name in the list of tag names, do the following
        for (TagName tagName : tagNames) {
			
            // Break the loop if the user clicks cancel
            if (progressPanel.getStatus() == ReportProgressPanel.ReportStatus.CANCELED) {
                break;
            }
            
            // Account for false user inputs
            if (ForensicReport_doc == null) {
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
                paragraphlist = new ArrayList<XWPFParagraph>();
                paragraphlist = ForensicReport_doc.getParagraphs();                

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
                    Content Content = tag.getContent();

                    // If the content object relating to this tagged file is an instance of AbstractFile class, do the following.
                    if (Content instanceof AbstractFile) {

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
                        if (null != ((AbstractFile) Content).getLocalAbsPath()) {
                            Path = ((AbstractFile) Content).getLocalAbsPath();                                
                        } else {
                            Path = tag.getContent().getUniquePath();                                
                        } 

                        // Retrieve MD5 Hash
                        md5hash = ((AbstractFile) Content).getMd5Hash();                                

                        // Retrieve Created Time
                        createdtime = ((AbstractFile) Content).getCtimeAsDate();   

                        // Retrieve Modified Time
                        modifiedtime = ((AbstractFile) Content).getMtimeAsDate();     

                        // Retrieve Accessed Time
                        accessedtime = ((AbstractFile) Content).getAtimeAsDate();  

                        // Retrieve the comment
                        if (tag.getComment() != null) {
                            comment = tag.getComment().trim();
                        }
                                                
                        // Retrieve the content, if the tagged file is an image.
                        for (String img_ext: img_exts) {
                            if (filename.contains(img_ext)) {
                                
                                // Create new file object as Dir, set to user home / .ForensicReportModule Directory
                                String dir = System.getProperty("user.home") + "\\.ForensicReportModule\\ImageFiles";
                                File Dir = new File(dir);

                                // If directory doesn't exist, create it
                                if (!Dir.exists()) {
                                    try{
                                        Dir.mkdir();
                                    } 
                                    catch(SecurityException se){
                                        Logger.getLogger(ForensicReportConfigPanel.class.getName()).log(Level.SEVERE, "Error creating folder " +Dir, se);
                                    }        
                                }

                                // Declare new file object, set to User home / .ForensicReportModule Directory + image           
                                File Image = new File(dir + "\\" +filename);
                                
                                // Write content to file & disk
                                try {
                                    ContentUtils.writeToFile(Content, Image);
                                } catch (IOException ex) {
                                    Logger.getLogger(ForensicReport.class.getName()).log(Level.SEVERE, "Error adding files", ex);
                                }
                                
                                // Create input stream of the file
                                try {
                                    image_is = new FileInputStream(Image);
                                } catch (FileNotFoundException ex) {
                                    Logger.getLogger(ForensicReport.class.getName()).log(Level.SEVERE, "Unable to create input stream", ex);
                                }
                            }
                        } //Finish retrieving image
                                               
                        // Count the amount of evidence headings found
                        heading_count = 0;

                        // Build the Table for this file with the retrieved metadata information
                        buildTables(tags, filename, Path, md5hash, comment, createdtime, modifiedtime, accessedtime, image_is);
                        
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
                Logger.getLogger(ForensicReport.class.getName()).log(Level.SEVERE, "Error adding files", ex);
                JOptionPane.showMessageDialog(null, "Error getting selected tags for case.", "File Export Error", JOptionPane.ERROR_MESSAGE);
            }            

        }
                        
        // If template 1 or 2 is selected, set the table colour and the table width of the existing table in the template to match the configured & generated tables.
        if (configPanel.Template_1_or_2_isSelected()) {
            tableRow = ForensicReport_doc.getTableArray(2).getRow(0);            
            for (int column=0; column<4; column++) {
                if (column==0) configureTable(tableRow, column, tableColour, "Item", fontColour, true, true);
                if (column==1) configureTable(tableRow, column, tableColour, "Serial Number", fontColour, true, true);
                if (column==2) configureTable(tableRow, column, tableColour, "Description", fontColour, true, false); 
                if (column==3) configureTable(tableRow, column, tableColour, "Type", fontColour, true, true);
            }
            // Set table width to 100%
            width = ForensicReport_doc.getTableArray(2).getCTTbl().addNewTblPr().addNewTblW();
            width.setType(STTblWidth.DXA);
            width.setW(BigInteger.valueOf(((6*1440)+938)));
        }
        
        // Write the Document in file system
        try {
            out = new FileOutputStream(new File(baseReportDir + "report." + file_extension));
        } catch(FileNotFoundException e){
            JOptionPane.showMessageDialog(null, "Unable to create new report.", "Create New Report Error", JOptionPane.ERROR_MESSAGE);
            Logger.getLogger(ForensicReportConfigPanel.class.getName()).log(Level.SEVERE, "Failed to create new report", e);
        }

        // Save the document to disk.                            
        if(out != null) {
            try {
                ForensicReport_doc.write(out);
                out.close();
            } catch(IOException e){
                JOptionPane.showMessageDialog(null, "Unable to save report.", "Save Report Error", JOptionPane.ERROR_MESSAGE);
                Logger.getLogger(ForensicReportConfigPanel.class.getName()).log(Level.SEVERE, "Failed to save report", e);
            }
        }
            
        // Manage the failed exports and display to user
        if (!failedExports.isEmpty()) {
            StringBuilder errorMessage = new StringBuilder("Failed to export the following files: ");
            for (int i=0; i<failedExports.size(); i++) {
                errorMessage.append(failedExports.get(i));
                if (failedExports.size()>1 && i<failedExports.size() - 1) {
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
            java.util.logging.Logger.getLogger(ForensicReport.class.getName()).log(Level.SEVERE, null, ex);
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
     * @param image_is
     */
    public void buildTables(List<ContentTag> tags, String filename, String Path, String md5hash, String comment, String createdtime, String modifiedtime, String accessedtime, FileInputStream image_is) 
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
                            tables[count] = ForensicReport_doc.insertNewTbl(cursor);
                        }

                        /**
                         * If this is not the first table created, set the cursor directly to after the previous
                         * created table, and create the table at this point.
                         * 
                         * Add the table to an array of tables which we have created
                         */                
                        else {
                            // Set cursor below the previous comment after previous table, if it exists
                            if(paragraph !=  null) {
                                cursor = paragraph.getCTP().newCursor();
                                cursor.toNextSibling();
                                tables[count] = ForensicReport_doc.insertNewTbl(cursor);
                            }
                            // If it doesn't exist, set the cursor to 2 siblings after previous table
                            else {
                                cursor = tables[(count-1)].getCTTbl().newCursor();
                                cursor.toNextSibling();
                                cursor.toNextSibling();
                                tables[count] = ForensicReport_doc.insertNewTbl(cursor);
                            }                                
                        }
                                                
//                        // Set table width to 100%, 1 inch = 1440                        
//                        width = tables[count].getCTTbl().addNewTblPr().addNewTblW();
//                        width.setType(STTblWidth.DXA);
//                        width.setW(BigInteger.valueOf((6*1440)+938));
//                        tables[count].getCTTbl().addNewTblGrid().addNewGridCol().setW(BigInteger.valueOf((1*1440)+85));
//                        tables[count].getCTTbl().getTblGrid().addNewGridCol().setW(BigInteger.valueOf((5*1440)+938-85));
                                                                       
                        // Create first row & 2nd column of table // File Name
                        XWPFTableRow tableRowOne = tables[count].getRow(0);
                        configureTable(tableRowOne, 0, tableColour, "File Name", fontColour, true, false);                        
                        tableRowOne.addNewTableCell();
                        if (filename != null) {
                            configureTable(tableRowOne, 1, "FFFFFF", filename, "000000", false, false);  
                        }
                        
                        // Set row 1 column 1 width
                        width = tables[count].getRow(0).getCell(0).getCTTc().addNewTcPr().addNewTcW();
                        width.setW(BigInteger.valueOf((1*1440)+85));
                        width.setType(STTblWidth.DXA);
                        
                        // Set row 1 column 2 width
                        width = tables[count].getRow(0).getCell(1).getCTTc().addNewTcPr().addNewTcW();
                        width.setW(BigInteger.valueOf((5*1440)+938-85));
                        width.setType(STTblWidth.DXA);
                                                                      
                        // Create second row of table // File Path
                        XWPFTableRow tableRowTwo = tables[count].createRow();
                        configureTable(tableRowTwo, 0, tableColour, "File Path", fontColour, true, false); 
                        if (Path != null) {
                            configureTable(tableRowTwo, 1, "FFFFFF", Path, "000000", false, false); 
                        }

                        // Create third row of table // Hash Value
                        XWPFTableRow tableRowThree = tables[count].createRow();
                        configureTable(tableRowThree, 0, tableColour, "Hash Value", fontColour, true, false); 
                        if (md5hash != null) {
                            configureTable(tableRowThree, 1, "FFFFFF", md5hash, "000000", false, false); 
                        }
                        else {
                            tableRowThree.getCell(1).setText("Hashes have not been calculated. Please configure and run an appropriate ingest module.");
                        }

                        // Create fourth row of table // Created time
                        XWPFTableRow tableRowFour = tables[count].createRow();
                        configureTable(tableRowFour, 0, tableColour, "Created time", fontColour, true, false);
                        if (Path != null) {
                            configureTable(tableRowFour, 1, "FFFFFF", createdtime, "000000", false, false); 
                        }
                        
                        // Set row 4 column 2 width
                        width = tables[count].getRow(3).getCell(1).getCTTc().addNewTcPr().addNewTcW();
                        width.setW(BigInteger.valueOf((((5*1440)+938-85) / 2) + 720));
                        width.setType(STTblWidth.DXA);
                        
//                        tableRowFour.addNewTableCell();

                        // Create fifth row of table // Modified time
                        XWPFTableRow tableRowFive = tables[count].createRow();
                        configureTable(tableRowFive, 0, tableColour, "Modified time", fontColour, true, false); 
                        if (Path != null) {
                            configureTable(tableRowFive, 1, "FFFFFF", modifiedtime, "000000", false, false); 
                        }
                        
                        // Set row 5 column 2 width
                        width = tables[count].getRow(4).getCell(1).getCTTc().addNewTcPr().addNewTcW();
                        width.setW(BigInteger.valueOf((((5*1440)+938-85) / 2) + 720));
                        width.setType(STTblWidth.DXA);

                        // Create sixth row of table // Accessed time
                        XWPFTableRow tableRowSix = tables[count].createRow();
                        configureTable(tableRowSix, 0, tableColour, "Accessed time", fontColour, true, false); 
                        if (Path != null) {
                            configureTable(tableRowSix, 1, "FFFFFF", accessedtime, "000000", false, false); 
                        }
                        
                        // Set row 6 column 2 width
                        width = tables[count].getRow(5).getCell(1).getCTTc().addNewTcPr().addNewTcW();
                        width.setW(BigInteger.valueOf((((5*1440)+938-85) / 2) + 720));
                        width.setType(STTblWidth.DXA);
                        
                        // Create third column // Content
                        if (image_is != null) {
                            //Create column
                        }

                        // Create paragraph after table // Comment
                        cursor = tables[count].getCTTbl().newCursor();
                        cursor.toNextSibling();
                        if (cursor != null) {
                            paragraph = ForensicReport_doc.insertNewParagraph(cursor);                                
                        }
                        if ( paragraph != null) {
                            if (comment != null && !(comment.isEmpty())) {                    
                                run = paragraph.createRun();
                                run.setText(comment);
                            }
                            if ((comment == null || comment.isEmpty()) && filename != null) {
                                run = paragraph.createRun();
                                run.setText("This table shows information about \"" +filename + "\"");
                            }  
                        }
                        
                        // Create gap before each table insert
                        cursor = tables[count].getCTTbl().newCursor();
                        if (cursor != null) {
                            paragraph2 = ForensicReport_doc.insertNewParagraph(cursor);                               
                        }
                        if (paragraph2 != null) {
                            run2 = paragraph2.createRun();
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
     * ConfigureTable Method.
     * Fourth Mutator Method.
     * 
     * Configures the font, text and styling of a row and column of a table.
     * 
     * @param row
     * @param column 
     */
    private void configureTable(XWPFTableRow row, int column, String tableColour, String title, String fontColour, boolean bold, boolean center) {
        
        // Set table colour accordingly
        row.getCell(column).setColor(tableColour);
        
        // Set text colour to black for lighter backgrounds
        if (tableColour.equals("00ffff") || tableColour.equals("ffff00")) {
            fontColour = "000000";
        }
                
        // Remove existing unchangeable paragraphs
        for (int x=0; x<row.getCell(column).getParagraphs().size(); x++) {
            row.getCell(column).removeParagraph(x);
        }

        // Add new paragraph
        paragraph = row.getCell(column).addParagraph();

        // Set and configure text of new paragraph accordingly
        run = paragraph.createRun();        
        setRun(run, "Calibri" , 10, fontColour, title, bold);
        
        // Set line spacing accordingly
        setSingleLineSpacing(paragraph);
        
        // Align text to the center accordingly
        if (center) {
            paragraph.setAlignment(ParagraphAlignment.CENTER);
        }
       
    }
    
    /**
     * SetRun Method.
     * Fifth Mutator Method.
     * 
     * Creates runs for configuring paragraphs.
     * 
     * @param run
     * @param fontFamily
     * @param fontSize
     * @param colorRGB
     * @param title
     * @param bold
     */
    private static void setRun (XWPFRun run, String fontFamily, int fontSize, String colorRGB, String title, boolean bold) {
        run.setFontFamily(fontFamily); 
        run.setFontSize(fontSize);
        run.setColor(colorRGB);
        run.setBold(bold);
        run.setText(title);
    }
    
    /**
     * SetSingleLineSpacing Method.
     * Sixth Mutator Method.
     * 
     * Configure the line spacing in paragraphs.
     * 
     * @param para 
     */
    public void setSingleLineSpacing(XWPFParagraph para) {        
        CTPPr ppr = para.getCTP().addNewPPr();
        CTSpacing spacing = ppr.addNewSpacing();
        spacing.setAfter(BigInteger.valueOf(0));
        spacing.setBefore(BigInteger.valueOf(0));        
        spacing.setLine(new BigInteger("240"));
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
        configPanel = new ForensicReportConfigPanel();
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
    public static synchronized ForensicReport getDefault() {
        if (instance == null) {
            instance = new ForensicReport();
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
    private XWPFRun run;
    private XWPFRun run2;
    private XWPFParagraph paragraph;
    private XWPFParagraph paragraph2;
    private XWPFTableRow tableRow;
    private CTTblWidth width;
    // End of variables declaration//GEN-END:variables 
}

