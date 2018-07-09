package Controlador;
import Modelo.ModeloExcel;
import Vista.Entrada;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.*;
import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;

public class ControladorExcel implements ActionListener{
    //Atributos:
    File archivo;
    int contAccion;
    
    //Instanciamos:
    ModeloExcel modeloE;
    Entrada entradaE;
    JFileChooser selecArchivo = new JFileChooser();

    //Controlador:
    public ControladorExcel() {
        this.modeloE = new ModeloExcel();
        this.entradaE = new Entrada();
        this.entradaE.setVisible(true);
        this.entradaE.setLocationRelativeTo(null);
        this.entradaE.jbImportar.addActionListener(this);
    }
    //Creamos un nuevo metodo:
    public void AgregarFiltro(){
        selecArchivo.setFileFilter(new FileNameExtensionFilter("Excel (*.xls)", "xls"));
        selecArchivo.setFileFilter(new FileNameExtensionFilter("Excel (*.xlsx)", "xlsx"));
    } 
    
    @Override
    public void actionPerformed(ActionEvent e) {
        contAccion++;
        if (contAccion == 1)AgregarFiltro();
        //Importar Excel:
        if (e.getSource() == entradaE.jbImportar) {
            if (selecArchivo.showDialog(null, "Seleccinar archivo") == JFileChooser.APPROVE_OPTION){
                archivo = selecArchivo.getSelectedFile();
                
                if(archivo.getName().endsWith("xls") || archivo.getName().endsWith("xlsx")){
                    JOptionPane.showMessageDialog(null, modeloE.Importar(archivo, entradaE.jTdatos)+ "\n Formato ."
                            + archivo.getName().substring(archivo.getName().lastIndexOf(".")+1),"IMPORTAR EXCEL", JOptionPane.INFORMATION_MESSAGE);
                }else{
                    JOptionPane.showMessageDialog(null, "Seleccione un formato valido");
                }
            }
        }
        //Exportar Excel:
       /* if (e.getSource() == entradaE.jbExportar) {
            if (selecArchivo.showDialog(null, "Exportar") == JFileChooser.APPROVE_OPTION){
                archivo = selecArchivo.getSelectedFile();
                
                if(archivo.getName().endsWith("xls") || archivo.getName().endsWith("xlsx")){
                    JOptionPane.showMessageDialog(null, modeloE.Exportar(archivo, entradaE.jtTablaEntrada)+ "\n Formato ."
                            + archivo.getName().substring(archivo.getName().lastIndexOf(".")+1));
                }else{
                    JOptionPane.showMessageDialog(null, "Seleccione un formato valido");
                }
            }
        }*/
    }
}
