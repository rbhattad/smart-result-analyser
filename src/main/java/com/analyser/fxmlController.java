package com.herewegoagain;

import java.io.File;
import java.io.IOException;
import javafx.event.ActionEvent;
import javafx.fxml.FXML;
import javafx.fxml.FXMLLoader;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.RadioButton;
import javafx.scene.text.Text;
import javafx.stage.FileChooser;
import javafx.stage.Modality;
import javafx.stage.Stage;

public class fxmlController {
    File filepdf;

    @FXML
    private Button selectfilebutton;

    @FXML
    private Button convertbutton;

    @FXML
    private RadioButton TE;

    @FXML
    private RadioButton BE;

    @FXML
    private RadioButton SE;

    @FXML
    private RadioButton BOTHSEM;

    @FXML
    private RadioButton SEM1;
    @FXML
    private Label labelfile;

    @FXML
    void convert(ActionEvent event) throws IOException {
        Converter A=new Converter();
        System.out.println("inside convert button");

        if(SE.isSelected() && !TE.isSelected() && !BE.isSelected())
        {
            System.out.println("SE selected\n\n\n\n\n\n\n");
            if(SEM1.isSelected() && !BOTHSEM.isSelected())
            {
                A.primarycontroller(filepdf, 2, 0);
            }
            else if(!SEM1.isSelected() && BOTHSEM.isSelected())
            {
                System.out.println("calling now");
                A.primarycontroller(filepdf, 2, 1); 
            }
        }
        else if(!SE.isSelected() && TE.isSelected() && !BE.isSelected())
        {
            if(SEM1.isSelected() && !BOTHSEM.isSelected())
            {
                A.primarycontroller(filepdf, 3, 0);
            }
            else if(!SEM1.isSelected() && BOTHSEM.isSelected())
            {
                A.primarycontroller(filepdf, 3, 1); 
            }
        }
        else if(!SE.isSelected() && !TE.isSelected() && BE.isSelected())
        {
             if(SEM1.isSelected() && !BOTHSEM.isSelected())
            {
                A.primarycontroller(filepdf, 4, 0);
            }
            else if(!SEM1.isSelected() && BOTHSEM.isSelected())
            {
                A.primarycontroller(filepdf, 4, 1);  
            }
        }
        System.out.println("task completed");
            
                 Stage dialog = new Stage();
                //dialog.initOwner(fxmlController);
                Scene dialogScene = new Scene(loadFXML("dialog"),200,150);
                dialog.setScene(dialogScene);
                dialog.show();
    }
    private static Parent loadFXML(String fxml) throws IOException {
        FXMLLoader fxmlLoader = new FXMLLoader(App.class.getResource(fxml + ".fxml"));
        return fxmlLoader.load();
    }
    @FXML
    void OKbutton(ActionEvent event)
    {
            ((Stage)(((Button)event.getSource()).getScene().getWindow())).close(); 
    }
    @FXML
    void selectfile(ActionEvent event) {
        
        FileChooser fc=new FileChooser();
         fc.getExtensionFilters().add(new FileChooser.ExtensionFilter("PDF Files","*.pdf"));
         filepdf =fc.showOpenDialog(null); 
              if(fc!=null){
                labelfile.setText(filepdf.getAbsolutePath());
            }
    }

}
