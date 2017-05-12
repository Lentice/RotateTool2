package lahome.rotateTool;

import javafx.application.Application;
import javafx.fxml.FXMLLoader;
import javafx.geometry.Bounds;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.scene.image.Image;
import javafx.stage.Stage;
import lahome.rotateTool.module.RotateCollection;
import lahome.rotateTool.view.RootController;

import java.util.prefs.Preferences;

public class Main extends Application {

    private Stage primaryStage;

    RotateCollection collection = new RotateCollection();
    RootController controller;

    private boolean maxStage = false;

    @Override
    public void start(Stage primaryStage) throws Exception {
        loadSetting();

        this.primaryStage = primaryStage;
        this.primaryStage.setMaximized(maxStage);

        FXMLLoader loader = new FXMLLoader();
        loader.setLocation(Main.class.getResource("view/Root.fxml"));
        Parent root = loader.load();
        this.primaryStage.setTitle("LA Rotate Tool 1.0.0");
        this.primaryStage.getIcons().add(
                new Image(Main.class.getResourceAsStream("/images/AppIcon.png")));
        this.primaryStage.setScene(new Scene(root));
        this.primaryStage.show();

        controller = loader.getController();
        controller.setMainApp(this);
    }

    @Override
    public void stop(){
        saveSetting();
    }

    private void loadSetting() {
        Preferences prefs = Preferences.userNodeForPackage(Main.class);

        maxStage = Boolean.valueOf(prefs.get("maxStage", "false"));
    }

    private void saveSetting() {
        Preferences prefs = Preferences.userNodeForPackage(Main.class);
        prefs.put("maxStage", String.valueOf(primaryStage.isMaximized()));
        controller.saveSetting();
    }


    public RotateCollection getCollection() {
        return collection;
    }

    public Stage getPrimaryStage() {
        return primaryStage;
    }

    public static void main(String[] args) {
        launch(args);
    }

}
