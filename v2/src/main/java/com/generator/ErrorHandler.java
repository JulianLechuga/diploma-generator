package com.generator;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import javax.swing.JFrame;
import javax.swing.JOptionPane;
import java.io.IOException;

public class ErrorHandler {

    private JFrame frame;

    public ErrorHandler(JFrame frame) {
        this.frame = frame;
    }

    public void handle(String message, Throwable exception) {
        if (exception instanceof IllegalArgumentException) {
            handleIllegalArgumentException((IllegalArgumentException) exception);
        } else if (exception instanceof IOException) {
            handleIOException((IOException) exception);
        } else if (exception instanceof InvalidFormatException) {
            handleInvalidFormatException((InvalidFormatException) exception);
        } else {
            handleGenericError(message, exception);
        }
    }

    private void handleIllegalArgumentException(IllegalArgumentException e) {
        JOptionPane.showMessageDialog(frame, "Error de argumento: " + e.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
    }

    private void handleIOException(IOException e) {
        JOptionPane.showMessageDialog(frame, "Error de lectura/escritura: " + e.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
    }

    private void handleInvalidFormatException(InvalidFormatException e) {
        JOptionPane.showMessageDialog(frame, "Formato inválido: " + e.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
    }

    private void handleGenericError(String message, Throwable exception) {
        JOptionPane.showMessageDialog(frame, message + "\n" + exception.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
    }
}
