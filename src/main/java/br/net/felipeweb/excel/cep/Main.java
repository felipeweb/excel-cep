package br.net.felipeweb.excel.cep;

import br.net.felipeweb.excel.cep.manipulator.ExcelManipulator;
import java.io.File;
import java.io.IOException;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import javafx.application.Application;
import javafx.geometry.Insets;
import javafx.scene.Scene;
import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.TextField;
import javafx.scene.layout.GridPane;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Pane;
import javafx.scene.layout.VBox;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

public class Main extends Application {


	@Override
	public void start(final Stage stage) {
		stage.setTitle("Excel Cep");

		final FileChooser fileChooser = new FileChooser();

		final Button openButton = new Button("Abrir Planilha");
		final Label label1 = new Label("Nome da planilha");
		final TextField sheetName = new TextField();
		HBox hbCollSheet = new HBox();
		hbCollSheet.getChildren().addAll(label1, sheetName);
		hbCollSheet.setSpacing(10);

		Label label2 = new Label("Letra da coluna do cep");
		TextField colCep = new TextField();
		HBox hbCollCep = new HBox();
		hbCollCep.getChildren().addAll(label2, colCep);
		hbCollCep.setSpacing(10);
		ExecutorService executorService = Executors.newFixedThreadPool(10);

		openButton.setOnAction(
				e -> {
					configureFileChooser(fileChooser);
					File file = fileChooser.showOpenDialog(stage);
					if (file != null && file.getName().endsWith(".xlsx")) {
						try {
							executorService.submit(() -> {
								try {
									new ExcelManipulator(file, sheetName.getText(), colCep.getText()).getAddress();
								} catch (IOException | InvalidFormatException e1) {
									throw new RuntimeException(e1);
								}
								System.exit(0);
							});

						} catch (Exception ex) {
							Alert alert = new Alert(AlertType.ERROR);
							alert.setTitle("ERRO");
							alert.setHeaderText("Erro inesperado");
							alert.showAndWait();
						}
					} else {
						Alert alert = new Alert(AlertType.ERROR);
						alert.setTitle("ERRO");
						alert.setHeaderText("Arquivo Invalido");
						alert.showAndWait();
					}
				});


		final GridPane fileGrid = new GridPane();

		GridPane.setConstraints(hbCollSheet, 0, 0);
		GridPane.setConstraints(hbCollCep, 0, 1);
		GridPane.setConstraints(openButton, 0, 2);
		fileGrid.setHgap(6);
		fileGrid.setVgap(6);
		fileGrid.getChildren().addAll(openButton);

		final Pane rootGroup = new VBox(12);
		rootGroup.getChildren().addAll(hbCollSheet, hbCollCep, fileGrid);
		rootGroup.setPadding(new Insets(12, 12, 12, 12));

		stage.setScene(new Scene(rootGroup, 500, 500));
		stage.show();
	}

	public static void main(String[] args) {
		Application.launch(args);
	}

	private static void configureFileChooser(final FileChooser fileChooser){
		fileChooser.setTitle("Abrir planilha");
		fileChooser.setInitialDirectory(new File(System.getProperty("user.home")));
	}

}
