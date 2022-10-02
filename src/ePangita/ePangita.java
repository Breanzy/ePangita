package EvalRecorder;

import java.awt.Color;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.DateFormatSymbols;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.filechooser.FileNameExtensionFilter;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

public class EvalRecorder extends javax.swing.JFrame {

    public EvalRecorder() {
        initComponents();
    }

    @SuppressWarnings("unchecked")
    
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jLabel4 = new javax.swing.JLabel();
        jPanel1 = new javax.swing.JPanel();
        MSULogo = new javax.swing.JLabel();
        Input = new javax.swing.JButton();
        UserName = new javax.swing.JTextField();
        IdNumber = new javax.swing.JTextField();
        ThesisName = new javax.swing.JTextField();
        Output = new javax.swing.JButton();
        jPanel2 = new javax.swing.JPanel();
        jLabel1 = new javax.swing.JLabel();
        CurrentFileAddress = new javax.swing.JLabel();
        jLabel2 = new javax.swing.JLabel();
        OJT = new javax.swing.JTextField();
        DateMonth = new com.toedter.calendar.JMonthChooser();
        DateYear = new com.toedter.calendar.JYearChooser();
        jPanel3 = new javax.swing.JPanel();
        DateLabel = new javax.swing.JLabel();

        jLabel4.setText("jLabel4");

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setBackground(new java.awt.Color(255, 0, 51));
        setForeground(new java.awt.Color(204, 0, 0));
        setResizable(false);

        jPanel1.setBackground(new java.awt.Color(255, 64, 64));

        MSULogo.setFont(new java.awt.Font("Uni Sans Thin CAPS", 1, 14)); // NOI18N
        MSULogo.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        MSULogo.setIcon(new javax.swing.ImageIcon(getClass().getResource("/EvalRecorder/seal-02 (1).png"))); // NOI18N
        MSULogo.setDebugGraphicsOptions(javax.swing.DebugGraphics.NONE_OPTION);
        MSULogo.setPreferredSize(new java.awt.Dimension(300, 305));

        Input.setFont(new java.awt.Font("Verdana", 1, 12)); // NOI18N
        Input.setForeground(new java.awt.Color(255, 51, 51));
        Input.setText("Select Raw Record File");
        Input.setBorder(new javax.swing.border.LineBorder(new java.awt.Color(255, 204, 0), 2, true));
        Input.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                InputActionPerformed(evt);
            }
        });

        UserName.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        UserName.setForeground(new java.awt.Color(153, 153, 153));
        UserName.setText("Name:");
        UserName.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(255, 204, 0), 2));
        UserName.addFocusListener(new java.awt.event.FocusAdapter() {
            public void focusGained(java.awt.event.FocusEvent evt) {
                UserNameFocusGained(evt);
            }
            public void focusLost(java.awt.event.FocusEvent evt) {
                UserNameFocusLost(evt);
            }
        });
        UserName.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                UserNameActionPerformed(evt);
            }
        });

        IdNumber.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        IdNumber.setForeground(new java.awt.Color(153, 153, 153));
        IdNumber.setText("ID Number:");
        IdNumber.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(255, 204, 0), 2));
        IdNumber.addFocusListener(new java.awt.event.FocusAdapter() {
            public void focusGained(java.awt.event.FocusEvent evt) {
                IdNumberFocusGained(evt);
            }
            public void focusLost(java.awt.event.FocusEvent evt) {
                IdNumberFocusLost(evt);
            }
        });
        IdNumber.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                IdNumberActionPerformed(evt);
            }
        });

        ThesisName.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        ThesisName.setForeground(new java.awt.Color(153, 153, 153));
        ThesisName.setText("Thesis Title (Optional):");
        ThesisName.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(255, 204, 0), 2));
        ThesisName.addFocusListener(new java.awt.event.FocusAdapter() {
            public void focusGained(java.awt.event.FocusEvent evt) {
                ThesisNameFocusGained(evt);
            }
            public void focusLost(java.awt.event.FocusEvent evt) {
                ThesisNameFocusLost(evt);
            }
        });
        ThesisName.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                ThesisNameActionPerformed(evt);
            }
        });

        Output.setFont(new java.awt.Font("Trebuchet MS", 1, 14)); // NOI18N
        Output.setForeground(new java.awt.Color(255, 51, 51));
        Output.setText("Create Evaluation Record");
        Output.setBorder(new javax.swing.border.LineBorder(new java.awt.Color(255, 204, 0), 2, true));
        Output.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                OutputActionPerformed(evt);
            }
        });

        jPanel2.setBackground(new java.awt.Color(255, 255, 255));
        jPanel2.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(255, 204, 0), 3));

        jLabel1.setBackground(new java.awt.Color(255, 255, 51));
        jLabel1.setFont(new java.awt.Font("MS Reference Sans Serif", 1, 18)); // NOI18N
        jLabel1.setForeground(new java.awt.Color(255, 51, 51));
        jLabel1.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        jLabel1.setText("MSU-IIT Evaluation Record Creator");
        jLabel1.setDebugGraphicsOptions(javax.swing.DebugGraphics.NONE_OPTION);

        javax.swing.GroupLayout jPanel2Layout = new javax.swing.GroupLayout(jPanel2);
        jPanel2.setLayout(jPanel2Layout);
        jPanel2Layout.setHorizontalGroup(
            jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel2Layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(jLabel1, javax.swing.GroupLayout.DEFAULT_SIZE, 354, Short.MAX_VALUE)
                .addGap(12, 12, 12))
        );
        jPanel2Layout.setVerticalGroup(
            jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel2Layout.createSequentialGroup()
                .addGap(0, 0, Short.MAX_VALUE)
                .addComponent(jLabel1, javax.swing.GroupLayout.PREFERRED_SIZE, 29, javax.swing.GroupLayout.PREFERRED_SIZE))
        );

        CurrentFileAddress.setFont(new java.awt.Font("Segoe UI", 0, 10)); // NOI18N
        CurrentFileAddress.setForeground(new java.awt.Color(255, 204, 0));
        CurrentFileAddress.setText("Current File Address: None");

        jLabel2.setForeground(new java.awt.Color(255, 204, 0));
        jLabel2.setText("Please fill in the required information below:");

        OJT.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        OJT.setForeground(new java.awt.Color(153, 153, 153));
        OJT.setText("Summer In-Plant Training (Optional):");
        OJT.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(255, 204, 0), 2));
        OJT.addFocusListener(new java.awt.event.FocusAdapter() {
            public void focusGained(java.awt.event.FocusEvent evt) {
                OJTFocusGained(evt);
            }
            public void focusLost(java.awt.event.FocusEvent evt) {
                OJTFocusLost(evt);
            }
        });
        OJT.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                OJTActionPerformed(evt);
            }
        });

        DateMonth.setDayChooser(null);
        DateMonth.setDoubleBuffered(false);
        DateMonth.setMonth(1);
        DateMonth.setName("DateMonth"); // NOI18N
        DateMonth.setOpaque(false);
        DateMonth.setRequestFocusEnabled(false);
        DateMonth.setVerifyInputWhenFocusTarget(false);
        DateMonth.setYearChooser(DateYear);

        DateYear.setDebugGraphicsOptions(javax.swing.DebugGraphics.NONE_OPTION);

        jPanel3.setBackground(new java.awt.Color(255, 255, 255));
        jPanel3.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(255, 204, 0), 2));
        jPanel3.setForeground(new java.awt.Color(255, 255, 255));

        DateLabel.setBackground(new java.awt.Color(255, 255, 255));
        DateLabel.setFont(new java.awt.Font("Segoe UI", 1, 13)); // NOI18N
        DateLabel.setForeground(new java.awt.Color(255, 51, 51));
        DateLabel.setText("Evaluation Date:");

        javax.swing.GroupLayout jPanel3Layout = new javax.swing.GroupLayout(jPanel3);
        jPanel3.setLayout(jPanel3Layout);
        jPanel3Layout.setHorizontalGroup(
            jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel3Layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(DateLabel, javax.swing.GroupLayout.DEFAULT_SIZE, 107, Short.MAX_VALUE)
                .addContainerGap())
        );
        jPanel3Layout.setVerticalGroup(
            jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(DateLabel, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
        );

        javax.swing.GroupLayout jPanel1Layout = new javax.swing.GroupLayout(jPanel1);
        jPanel1.setLayout(jPanel1Layout);
        jPanel1Layout.setHorizontalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jPanel2, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addGroup(jPanel1Layout.createSequentialGroup()
                        .addGap(12, 12, 12)
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(CurrentFileAddress, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                            .addGroup(jPanel1Layout.createSequentialGroup()
                                .addComponent(Input, javax.swing.GroupLayout.PREFERRED_SIZE, 163, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                .addComponent(Output, javax.swing.GroupLayout.PREFERRED_SIZE, 182, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addGap(8, 8, 8))
                            .addGroup(jPanel1Layout.createSequentialGroup()
                                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                    .addComponent(UserName, javax.swing.GroupLayout.PREFERRED_SIZE, 366, javax.swing.GroupLayout.PREFERRED_SIZE)
                                    .addComponent(IdNumber, javax.swing.GroupLayout.PREFERRED_SIZE, 366, javax.swing.GroupLayout.PREFERRED_SIZE)
                                    .addComponent(jLabel2, javax.swing.GroupLayout.PREFERRED_SIZE, 286, javax.swing.GroupLayout.PREFERRED_SIZE))
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, 54, Short.MAX_VALUE)
                                .addComponent(MSULogo, javax.swing.GroupLayout.PREFERRED_SIZE, 99, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addGap(47, 47, 47))
                            .addGroup(jPanel1Layout.createSequentialGroup()
                                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                    .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                                        .addComponent(OJT)
                                        .addComponent(ThesisName, javax.swing.GroupLayout.PREFERRED_SIZE, 364, javax.swing.GroupLayout.PREFERRED_SIZE))
                                    .addGroup(jPanel1Layout.createSequentialGroup()
                                        .addComponent(jPanel3, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                        .addComponent(DateMonth, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                        .addComponent(DateYear, javax.swing.GroupLayout.PREFERRED_SIZE, 105, javax.swing.GroupLayout.PREFERRED_SIZE)))
                                .addGap(0, 0, Short.MAX_VALUE)))))
                .addContainerGap())
        );
        jPanel1Layout.setVerticalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                .addComponent(jPanel2, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(jPanel1Layout.createSequentialGroup()
                        .addComponent(jLabel2)
                        .addGap(18, 18, 18)
                        .addComponent(UserName, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                        .addComponent(IdNumber, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                    .addComponent(MSULogo, javax.swing.GroupLayout.PREFERRED_SIZE, 97, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                    .addComponent(jPanel3, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addComponent(DateYear, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addComponent(DateMonth, javax.swing.GroupLayout.PREFERRED_SIZE, 29, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(7, 7, 7)
                .addComponent(ThesisName, javax.swing.GroupLayout.PREFERRED_SIZE, 29, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(OJT, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(Output, javax.swing.GroupLayout.PREFERRED_SIZE, 34, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(Input, javax.swing.GroupLayout.PREFERRED_SIZE, 30, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(CurrentFileAddress, javax.swing.GroupLayout.PREFERRED_SIZE, 21, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(16, 16, 16))
        );

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jPanel1, javax.swing.GroupLayout.Alignment.TRAILING, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jPanel1, javax.swing.GroupLayout.Alignment.TRAILING, javax.swing.GroupLayout.PREFERRED_SIZE, 331, javax.swing.GroupLayout.PREFERRED_SIZE)
        );

        pack();
    }// </editor-fold>//GEN-END:initComponents

    private void InputActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_InputActionPerformed
        
        //Activates when the button for selecting Raw Files is pressed
        if(evt.getSource() == Input)
        {
            JFileChooser loadFile = new JFileChooser();
            FileNameExtensionFilter filter = new FileNameExtensionFilter("XLSX Files", "xlsx");
            
            //allows user to select the location of the raw file
            loadFile.setFileFilter(filter);
            loadFile.setDialogTitle("Select Raw Record File");
            loadFile.setSelectedFile(new File("Raw Record.xlsx"));
            
            if(loadFile.showSaveDialog(null) == JFileChooser.APPROVE_OPTION)
            {
                Address = loadFile.getSelectedFile().getAbsolutePath();
                JOptionPane.showMessageDialog(this, "Successfully selected " + Address + " as Raw Record File!");
                CurrentFileAddress.setText("Current File Address: " + Address);
            }
        }
    }//GEN-LAST:event_InputActionPerformed

    private void OutputActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_OutputActionPerformed
        
        //Activates when the button for creating new Eval Record is pressed
        
        //Checks to see if all important informations are present
        if (UserName.getText().equals("Name:")) JOptionPane.showMessageDialog(null, "Name Required!" ); 
        else if(IdNumber.getText().equals("ID Number:")) JOptionPane.showMessageDialog(null, "ID Number Required!" );
        else if(Address == null) JOptionPane.showMessageDialog(null, "No Raw Record File selected!");
   
        else
        {
            //initializes all received values
            name = UserName.getText();
            idNumber = IdNumber.getText();
            
            date =  new DateFormatSymbols().getMonths()[DateMonth.getMonth()] + " " + DateYear.getYear();
            
            if(!ThesisName.getText().equals("Thesis Title (Optional):"))
                thesisName = ThesisName.getText();
            
            if(!OJT.getText().equals("Summer In-Plant Training (Optional):"))
                ojt = OJT.getText();
            
            CreateOutputFile();   
        }
        
    }//GEN-LAST:event_OutputActionPerformed

    private void UserNameActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_UserNameActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_UserNameActionPerformed

    private void UserNameFocusGained(java.awt.event.FocusEvent evt) {//GEN-FIRST:event_UserNameFocusGained
        // TODO add your handling code here:
        if(UserName.getText().equals("Name:"))
        {
            UserName.setText("");
            UserName.setForeground(new Color(0, 0, 0));
        }
    }//GEN-LAST:event_UserNameFocusGained

    private void UserNameFocusLost(java.awt.event.FocusEvent evt) {//GEN-FIRST:event_UserNameFocusLost
        // TODO add your handling code here:
        if(UserName.getText().equals(""))
        {
            UserName.setText("Name:");
            UserName.setForeground(new Color(153, 153, 153));
        }
    }//GEN-LAST:event_UserNameFocusLost

    private void ThesisNameFocusGained(java.awt.event.FocusEvent evt) {//GEN-FIRST:event_ThesisNameFocusGained
        // TODO add your handling code here:
        if(ThesisName.getText().equals("Thesis Title (Optional):"))
        {
            ThesisName.setText("");
            ThesisName.setForeground(new Color(0, 0, 0));
        }
    }//GEN-LAST:event_ThesisNameFocusGained

    private void ThesisNameFocusLost(java.awt.event.FocusEvent evt) {//GEN-FIRST:event_ThesisNameFocusLost
        // TODO add your handling code here:
        
        if(ThesisName.getText().equals(""))
        {
            ThesisName.setText("Thesis Title (Optional):");
            ThesisName.setForeground(new Color(153, 153, 153));
        }
    }//GEN-LAST:event_ThesisNameFocusLost

    private void ThesisNameActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_ThesisNameActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_ThesisNameActionPerformed

    private void IdNumberFocusGained(java.awt.event.FocusEvent evt) {//GEN-FIRST:event_IdNumberFocusGained
        // TODO add your handling code here:
        if(IdNumber.getText().equals("ID Number:"))
        {
            IdNumber.setText("");
            IdNumber.setForeground(new Color(0, 0, 0));
        }
    }//GEN-LAST:event_IdNumberFocusGained
                                 
    private void IdNumberFocusLost(java.awt.event.FocusEvent evt) {//GEN-FIRST:event_IdNumberFocusLost
        // TODO add your handling code here:
        if(IdNumber.getText().equals(""))
        {
            IdNumber.setText("ID Number:");
            IdNumber.setForeground(new Color(153, 153, 153));
        }
    }//GEN-LAST:event_IdNumberFocusLost

    private void IdNumberActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_IdNumberActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_IdNumberActionPerformed

    private void OJTFocusGained(java.awt.event.FocusEvent evt) {//GEN-FIRST:event_OJTFocusGained
        // TODO add your handling code here:
        if(OJT.getText().equals("Summer In-Plant Training (Optional):"))
        {
            OJT.setText("");
            OJT.setForeground(new Color(0, 0, 0));
        }
    }//GEN-LAST:event_OJTFocusGained

    private void OJTFocusLost(java.awt.event.FocusEvent evt) {//GEN-FIRST:event_OJTFocusLost
        // TODO add your handling code here:
        if(OJT.getText().equals(""))
        {
            OJT.setText("Summer In-Plant Training (Optional):");
            OJT.setForeground(new Color(153, 153, 153));
        }
    }//GEN-LAST:event_OJTFocusLost

    private void OJTActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_OJTActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_OJTActionPerformed
    
    //String values of user inputted data for New Evaluation File
    public String Address, thesisName, idNumber, date, name, ojt;
    
    private void CreateOutputFile()
    {
        
        //Initializing Course List
        List<Course> courseList = Arrays.asList(new Course("GEC", 11, 3, 9, 53), new Course("FIL", 22, 3, 2, 54), 
        new Course("FPE", 24, 3, 1, 54),  new Course("HIS", 25, 3, 1, 54), new Course("CHM", 28, 3, 2, 55), new Course("PHY", 30, 3, 2, 55),
        new Course("MAT", 34, 3, 3, 56), new Course("ENS", 39, 3, 8, 57), new Course("EEE", 49, 3, 12, 58), new Course("COE", 11, 9, 27, 60), 
        new Course("PED", 40, 9, 4, 61), new Course("NST", 47, 9, 2, 62), new Course("ELEC", 63, 3, 3, 59));

        //Creates a file stream to access the data of the raw file
        FileInputStream inputFile;
        XSSFWorkbook inputWorkbook = new XSSFWorkbook();

        try{
            inputFile = new FileInputStream(new File(Address));
            inputWorkbook = new XSSFWorkbook(inputFile);
        } catch (FileNotFoundException ex) { JOptionPane.showMessageDialog(this, "Raw File Location Invalid");} 
        catch (IOException ex) { JOptionPane.showMessageDialog(this, "Unexceted error occured. Raw File might be corrupted or damaged");}
        
        //create inputSheet (to access raw file's Sheets)
        XSSFSheet inputSheet = inputWorkbook.getSheet("Sheet1");
        
        //6 Is the index of the first column of subject entries in the raw file. The program starts recording data here
        int cellColumn = 6;
        
        //create inputRow (to access raw file's row cells)
        XSSFRow inputRow = inputSheet.getRow(cellColumn);
        
        //records all the data of the row entries
        while(inputRow != null)
        {  
            //Gets the value of the cell in the third row (#2), which is the Subject Code (Ex: PED001),
            //then splits string value to get the course name and the subject number (PED, 001)
            String splitSubCode[] = inputRow.getCell(2).getStringCellValue().split("(?<=\\D)(?=\\d)", 2);
            //then gets subject grade in its cell row address (#7) 
            String subGrade = new DataFormatter().formatCellValue(inputRow.getCell(7));
            
            //boolean to check if course code is known
            boolean isKnownCourse = false;
            
            for(Course course : courseList)
            {
                //check which course it belongs
                if(course.courseName.equals(splitSubCode[0]))
                {
                    course.subList.add(new Subject(Double.parseDouble(splitSubCode[1]), subGrade));
                    isKnownCourse = true;
                    break;
                }
            }
            //no Known Course implies a course that was not initialized. It is then assumed as an ELECTIVE subject
            if(!isKnownCourse)
                courseList.get(courseList.size() - 1).subList.add(new Subject(splitSubCode[0], Double.parseDouble(splitSubCode[1]), subGrade));     
            
            //shifts to the next column
            cellColumn++;
            inputRow = inputSheet.getRow(cellColumn);
        }
        
        //creates a new file stream to access a template output file (Test.xlsx), where the final output of the eval record will be based from
        FileInputStream templateFile;
        XSSFWorkbook templateWorkbook = new XSSFWorkbook();
          
        try
        {
          templateFile = new FileInputStream(new File(System.getProperty("user.dir") + "\\Test.xlsx"));
          templateWorkbook = new XSSFWorkbook(templateFile);
        } 
        catch (FileNotFoundException ex){ JOptionPane.showMessageDialog(this, "Template File not found"); } 
        catch (IOException ex) { JOptionPane.showMessageDialog(this, "Unexceted error occured. Template File might be corrupted or damaged"); }
        
        XSSFSheet templateSheet = templateWorkbook.getSheet("Sheet1");
        
        //Creates a Bold font format to modify font property of OJT cell
        Font newFont = templateWorkbook.createFont();
        newFont.setBold(true);
        
        applyFontStyle(ojt, 71, 0, templateSheet, newFont, 26);
        
        //Adds an additional Underlined font format to modify Name, Thesis Title, Id Number, and Evaluation Date cells
        newFont.setUnderline(Font.U_SINGLE);

        applyFontStyle(name, 6, 0, templateSheet, newFont, 17);
        applyFontStyle(idNumber, 6, 5, templateSheet, newFont, 7);
        applyFontStyle(date, 6, 7, templateSheet, newFont, 18);
        applyFontStyle(thesisName, 7, 0, templateSheet, newFont, 22);

        //adds grades in the template file, based on information from the raw file
        for(Course course : courseList)
        {
            for (int i = course.courseRowAddress; i < (course.courseRowAddress + course.totalSubject); i++)
            {
                XSSFRow row = templateSheet.getRow(i); 
                for(Subject subject : course.subList)
                {
                    
                    //Special condition for ELEC subjects, where their unique Course code will also be printed
                    if(course.courseName.equals("ELEC"))
                    {
                        row.getCell(course.courseGradeAddress - 3).setCellValue(subject.courseName);
                        row.getCell(course.courseGradeAddress - 2).setCellValue(subject.subNumber);
                    }
                    
                    //checks if the subject number in the template file matches with any of the subject numbers in the subList
                    if(subject.subNumber == row.getCell(course.courseGradeAddress - 2).getNumericCellValue() || course.courseName.equals("ELEC"))
                    {
                        //matched number but failed grade means it will not be printed and be included in unit count
                        if(subject.subGrade.equals("5")) 
                            break;
                        
                        //Prints grade as a Double number
                        else if(!subject.subGrade.equals("INC") && !subject.subGrade.equals("P") )
                            row.getCell(course.courseGradeAddress).setCellValue(Double.parseDouble(subject.subGrade));    
                        //Prints grade as a String (for "INC" or "P" grades)
                        else
                            row.getCell(course.courseGradeAddress).setCellValue(subject.subGrade);
                          
                        //Adds unit to total unit of the course
                        int newTotalUnit = (int)row.getCell(course.courseGradeAddress + 1).getNumericCellValue();   
                        
                        try {
                            int oldTotalUnit = (int)templateSheet.getRow(course.totalUnitAddress).getCell(10).getNumericCellValue();
                            templateSheet.getRow(course.totalUnitAddress).getCell(10).setCellValue(newTotalUnit + oldTotalUnit);
                        } catch(IllegalStateException noInitialValue) {
                        templateSheet.getRow(course.totalUnitAddress).createCell(10).setCellValue(newTotalUnit); }
                        
                        course.subList.remove(subject);
                        break; 
                    }
                }
            }
            
            //Checks if there are still subjects left unassigned. These unassigned subjects are assumed as ELECTIVE subjects.
            if(!course.subList.isEmpty())
                for(Subject subject : course.subList)
                    courseList.get(courseList.size() - 1).subList.add(new Subject(course.courseName, subject.subNumber, subject.subGrade));                  
        }

        //Adds and prints the total units earned from all course units
        int totalUnits = 0;
        for(int a = 53; a < 62; a++) { totalUnits += (int)templateSheet.getRow(a).getCell(10).getNumericCellValue(); }
        templateSheet.getRow(63).getCell(10).setCellValue(totalUnits);
        
        //Creates a new Evaluation Record File from the template file
        JFileChooser saveFile = new JFileChooser();
        saveFile.setDialogTitle("Save New Evaluation Record File: ");
        saveFile.setSelectedFile(new File("New Eval Record.xlsx"));
        if(saveFile.showSaveDialog(null) == JFileChooser.APPROVE_OPTION)
        {
            File output = saveFile.getSelectedFile();
            try(FileOutputStream out = new FileOutputStream(output))
            {
                templateWorkbook.write(out);
                out.close();
            } catch (FileNotFoundException ex) {
                Logger.getLogger(EvalRecorder.class.getName()).log(Level.SEVERE, null, ex);
            } catch (IOException ex) {
                Logger.getLogger(EvalRecorder.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
        
        JOptionPane.showMessageDialog(this, "File creation complete!");
    }
    
    public void applyFontStyle(String value, int row, int cell, XSSFSheet sheet, Font font, int textLimit)
    {
        XSSFRichTextString newRTS = new XSSFRichTextString(sheet.getRow(row).getCell(cell).getStringCellValue() + value);
        newRTS.applyFont(textLimit, newRTS.length(), font); 
        sheet.getRow(row).getCell(cell).setCellValue(newRTS);
    }
    public static class Course
    {
        private String courseName;
        ArrayList<Subject> subList;
        private int totalSubject;
   
        //These values are from cell addresses of the file's final output.
        private int courseRowAddress;
        private int courseUnitAddress = courseRowAddress + 1;
        private int courseGradeAddress;
        private int totalUnitAddress;
        
        public Course(String courseName, int courseRowAddress, int courseGradeAddress, int totalSubject, int totalUnitAddress)
        { 
            subList = new ArrayList<>();
            this.courseName = courseName; 
            this.courseRowAddress = courseRowAddress; 
            this.courseGradeAddress = courseGradeAddress; 
            this.totalSubject = totalSubject;
            this.totalUnitAddress = totalUnitAddress;
        }              
    }
   
    public static class Subject
    {
        private double subNumber;
        private String subGrade;
        
        //exclusive variable for ELECTIVE subjects, as their course code differs from the rest
        private String courseName;
                
        public Subject(double subNumber, String subGrade){ this.subNumber = subNumber; this.subGrade = subGrade; }     
        public Subject(String courseName, double subNumber, String subGrade){ this.courseName = courseName; this.subNumber = subNumber; this.subGrade = subGrade; }   
    }
    
    public static void main(String args[]) {
        /* Set the Nimbus look and feel */
        //<editor-fold defaultstate="collapsed" desc=" Look and feel setting code (optional) ">
        /* If Nimbus (introduced in Java SE 6) is not available, stay with the default look and feel.
         * For details see http://download.oracle.com/javase/tutorial/uiswing/lookandfeel/plaf.html 
         */
        try {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
                if ("Nimbus".equals(info.getName())) {
                    javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (ClassNotFoundException | InstantiationException | IllegalAccessException | javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(EvalRecorder.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>
        //</editor-fold>
        
        //</editor-fold>
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(() -> {
            new EvalRecorder().setVisible(true);
        });       
    }
    
    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JLabel CurrentFileAddress;
    private javax.swing.JLabel DateLabel;
    private com.toedter.calendar.JMonthChooser DateMonth;
    private com.toedter.calendar.JYearChooser DateYear;
    private javax.swing.JTextField IdNumber;
    private javax.swing.JButton Input;
    private javax.swing.JLabel MSULogo;
    private javax.swing.JTextField OJT;
    private javax.swing.JButton Output;
    private javax.swing.JTextField ThesisName;
    private javax.swing.JTextField UserName;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JLabel jLabel4;
    private javax.swing.JPanel jPanel1;
    private javax.swing.JPanel jPanel2;
    private javax.swing.JPanel jPanel3;
    // End of variables declaration//GEN-END:variables
}
