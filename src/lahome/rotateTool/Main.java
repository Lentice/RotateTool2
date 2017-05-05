package lahome.rotateTool;

import javafx.application.Application;
import javafx.fxml.FXMLLoader;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.scene.image.Image;
import javafx.stage.Stage;
import lahome.rotateTool.module.RotateCollection;
import lahome.rotateTool.view.RootController;

public class Main extends Application {

    private Stage primaryStage;

    RotateCollection collection = new RotateCollection();
    RootController controller;

    @Override
    public void start(Stage primaryStage) throws Exception {
        this.primaryStage = primaryStage;

        FXMLLoader loader = new FXMLLoader();
        loader.setLocation(Main.class.getResource("view/Root.fxml"));
        Parent root = loader.load();
        this.primaryStage.setTitle("Hello World");
        this.primaryStage.getIcons().add(
                new Image("file:resources/images/AppIcon.png"));
        this.primaryStage.setScene(new Scene(root));
        this.primaryStage.show();

        controller = loader.getController();
        controller.setMainApp(this);

        controller.loadSetting();
    }

    @Override
    public void stop(){
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
