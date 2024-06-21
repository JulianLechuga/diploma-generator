package com.generator;

import org.apache.poi.xwpf.usermodel.*;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.*;

public class DiplomaGenerator {
    private JTextField templateField;
    private JTextField namesField;
    private JLabel previewLabel;
    private JComboBox<String> fontSelector;
    private JFrame frame;
    private ErrorHandler errorHandler;

    public DiplomaGenerator() {
        frame = new JFrame("Generador de Certificados");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setSize(800, 600);

        errorHandler = new ErrorHandler(frame); // Inicializar el ErrorHandler

        // Panel principal con GridBagLayout
        JPanel mainPanel = new JPanel();
        mainPanel.setLayout(new GridBagLayout());
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.gridx = 0;
        gbc.gridy = 0;
        gbc.anchor = GridBagConstraints.WEST;
        gbc.insets = new Insets(10, 10, 10, 10); // Margen

        // Etiqueta y campo de Plantilla
        JLabel templateLabel = new JLabel("Plantilla:");
        mainPanel.add(templateLabel, gbc);

        gbc.gridx++;
        templateField = new JTextField(30);
        mainPanel.add(templateField, gbc);

        gbc.gridx++;
        JButton templateButton = new JButton("Seleccionar");
        templateButton.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                chooseTemplate();
            }
        });
        mainPanel.add(templateButton, gbc);

        // Nueva fila para Archivo de Nombres
        gbc.gridx = 0;
        gbc.gridy++;
        JLabel namesLabel = new JLabel("Archivo de Nombres:");
        mainPanel.add(namesLabel, gbc);

        gbc.gridx++;
        namesField = new JTextField(30);
        mainPanel.add(namesField, gbc);

        gbc.gridx++;
        JButton namesButton = new JButton("Seleccionar");
        namesButton.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                chooseNamesFile();
            }
        });
        mainPanel.add(namesButton, gbc);

        // Nueva fila para selector de fuente
        gbc.gridx = 0;
        gbc.gridy++;
        JLabel fontLabel = new JLabel("Fuente para el nombre:");
        mainPanel.add(fontLabel, gbc);

        gbc.gridx++;
        fontSelector = new JComboBox<>(GraphicsEnvironment.getLocalGraphicsEnvironment().getAvailableFontFamilyNames());
        fontSelector.setSelectedItem("Edwardian Script ITC");
        mainPanel.add(fontSelector, gbc);

        // Botón para generar certificados
        gbc.gridx = 0;
        gbc.gridy++;
        gbc.gridwidth = 3;
        gbc.fill = GridBagConstraints.HORIZONTAL;
        JButton generateButton = new JButton("Generar Certificados");
        generateButton.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                generateCertificates();
            }
        });
        mainPanel.add(generateButton, gbc);

        // Vista previa
        gbc.gridy++;
        gbc.gridx = 0;
        gbc.gridwidth = 3;
        gbc.fill = GridBagConstraints.BOTH;
        previewLabel = new JLabel();
        previewLabel.setBorder(BorderFactory.createLineBorder(Color.GRAY));
        previewLabel.setPreferredSize(new Dimension(600, 200)); // Tamaño deseado para la vista previa
        mainPanel.add(previewLabel, gbc);

        // Agregar el panel principal al frame
        frame.add(mainPanel);
        frame.setVisible(true);
    }

    private void chooseTemplate() {
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setFileFilter(new FileNameExtensionFilter("Documentos Word", "docx"));
        if (fileChooser.showOpenDialog(frame) == JFileChooser.APPROVE_OPTION) {
            templateField.setText(fileChooser.getSelectedFile().getAbsolutePath());
            updatePreview();
        }
    }

    private void chooseNamesFile() {
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setFileFilter(new FileNameExtensionFilter("Archivos de Texto", "txt"));
        if (fileChooser.showOpenDialog(frame) == JFileChooser.APPROVE_OPTION) {
            namesField.setText(fileChooser.getSelectedFile().getAbsolutePath());
            updatePreview();
        }
    }

    private void generateCertificates() {
        String templatePath = templateField.getText();
        String namesPath = namesField.getText();

        if (templatePath.isEmpty() || namesPath.isEmpty()) {
            errorHandler.handle("Por favor, seleccione tanto la plantilla como el archivo de nombres.", new IllegalArgumentException("Plantilla o archivo de nombres vacío."));
            return;
        }

        try (BufferedReader reader = new BufferedReader(new FileReader(namesPath))) {
            String name;
            while ((name = reader.readLine()) != null) {
                createCertificate(templatePath, name);
            }
            JOptionPane.showMessageDialog(frame, "Documentos creados exitosamente.");
        } catch (IOException e) {
            errorHandler.handle("Error al generar certificados.", e);
        }
    }

    private void createCertificate(String templatePath, String name) {
        try {
            FileInputStream fis = new FileInputStream(templatePath);
            XWPFDocument document = new XWPFDocument(fis);

            String selectedFont = (String) fontSelector.getSelectedItem();

            for (XWPFParagraph paragraph : document.getParagraphs()) {
                for (XWPFRun run : paragraph.getRuns()) {
                    String text = run.getText(0);
                    if (text != null && text.contains("[NOMBRE]")) {
                        text = text.replace("[NOMBRE]", name);
                        run.setText(text, 0);
                        run.setFontFamily(selectedFont);
                        run.setFontSize(28);
                    }
                }
            }

            FileOutputStream fos = new FileOutputStream("Diploma_" + name.replace(" ", "_") + ".docx");
            document.write(fos);
            fos.close();
            document.close();
        } catch (IOException e) {
            errorHandler.handle("Error al crear el certificado para: " + name, e);
        }
    }

    private void updatePreview() {
        String templatePath = templateField.getText();
        String namesPath = namesField.getText();

        if (templatePath.isEmpty() || namesPath.isEmpty()) {
            errorHandler.handle("Por favor, seleccione tanto la plantilla como el archivo de nombres.", new IllegalArgumentException("Plantilla o archivo de nombres vacío."));
            return;
        }

        try (BufferedReader reader = new BufferedReader(new FileReader(namesPath))) {
            String firstLine = reader.readLine(); // Leer solo la primera línea
            if (firstLine != null && !firstLine.isEmpty()) {
                // Usar la primera línea completa
                String firstName = firstLine;

                // Crear documento temporal
                FileInputStream fis = new FileInputStream(templatePath);
                XWPFDocument document = new XWPFDocument(fis);

                String selectedFont = (String) fontSelector.getSelectedItem();

                // Reemplazar [NOMBRE] con el nombre completo de la primera línea
                for (XWPFParagraph paragraph : document.getParagraphs()) {
                    for (XWPFRun run : paragraph.getRuns()) {
                        String text = run.getText(0);
                        if (text != null && text.contains("[NOMBRE]")) {
                            text = text.replace("[NOMBRE]", firstName);
                            run.setText(text, 0);
                            run.setFontFamily(selectedFont);
                            run.setFontSize(28);
                        }
                    }
                }

                // Crear un archivo temporal para escribir el documento
                File tempFile = File.createTempFile("tempPreview", ".docx");
                FileOutputStream fos = new FileOutputStream(tempFile);
                document.write(fos);

                // Cerrar flujos
                fos.close();
                fis.close();
                document.close();

                // Leer el archivo temporal y mostrarlo en un JTextArea
                String previewText = readDocumentToString(tempFile);

                // Mostrar la vista previa en un JTextArea
                JTextArea previewTextArea = new JTextArea(previewText);
                previewTextArea.setEditable(false);

                // Mostrar la vista previa en un JOptionPane
                JOptionPane.showMessageDialog(frame, new JScrollPane(previewTextArea), "Vista Previa", JOptionPane.PLAIN_MESSAGE);

                // Borrar el archivo temporal después de mostrar la vista previa
                tempFile.delete();
            } else {
                JOptionPane.showMessageDialog(frame, "El archivo de nombres está vacío.");
            }
        } catch (Exception e) {
            errorHandler.handle("Error al generar vista previa.", e);
        }
    }

    private String readDocumentToString(File file) throws IOException {
        FileInputStream fis = new FileInputStream(file);
        XWPFDocument document = new XWPFDocument(fis);

        StringWriter stringWriter = new StringWriter();
        for (XWPFParagraph paragraph : document.getParagraphs()) {
            stringWriter.write(paragraph.getText());
            stringWriter.write("\n");
        }

        fis.close();
        document.close();

        return stringWriter.toString();
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> {
            new DiplomaGenerator();
        });
    }
}
