package com.example;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;

public class DocToHtmlConverter {

    public static void main(String[] args) {
        // Replace these paths with your input DOC file and desired output HTML file
        String inputDocPath = "C:\\Users\\Balakrishnan\\Documents\\test\\input.docx";
        String outputHtmlPath = "C:\\Users\\Balakrishnan\\Documents\\test\\output.html";

        try {
            // Build the Pandoc command with additional options for table borders
            ProcessBuilder processBuilder = new ProcessBuilder(
                    "C:\\Users\\Balakrishnan\\Downloads\\pandoc-3.1.6.1-windows-x86_64\\pandoc-3.1.6.1\\pandoc.exe",
                    inputDocPath,
                    "-o",
                    outputHtmlPath,
                    "--standalone",      // Include a table of contents
                    "--toc-depth=3", // Set TOC depth
                    "--css=C:\\Users\\Balakrishnan\\Downloads\\docx.converters-1.0.4-sample\\docx.converters-1.0.4\\src\\fr\\opensagres\\xdocreport\\samples\\docx\\converters\\xhtml\\style.css" // Apply custom CSS to style tables
            );

            // Redirect error stream to output stream
            processBuilder.redirectErrorStream(true);

            // Start the Pandoc process
            Process process = processBuilder.start();

            // Read and print Pandoc's output (optional)
            BufferedReader reader = new BufferedReader(new InputStreamReader(process.getInputStream()));
            String line;
            while ((line = reader.readLine()) != null) {
                System.out.println(line);
            }

            // Wait for the process to finish
            int exitCode = process.waitFor();
            if (exitCode == 0) {
                System.out.println("Conversion completed successfully.");
            } else {
                System.err.println("Error occurred during conversion.");
            }
        } catch (IOException | InterruptedException e) {
            System.out.println(e);
        }
    }
}
