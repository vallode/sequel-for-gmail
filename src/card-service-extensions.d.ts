import "google-apps-script";

declare global {
  namespace GoogleAppsScript {
    namespace Card_Service {
      interface TextButton {
        setMaterialIcon(icon: Card_Service.MaterialIcon): TextButton;
      }
    }
  }
}
