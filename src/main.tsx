import ReactDOM from "react-dom/client";
import App from "./App";
import PowerProvider from "./assets/PowerProvider";
import "./style.css";

ReactDOM.createRoot(document.getElementById("app")!).render(
  <PowerProvider>
    <App />
  </PowerProvider>
);
