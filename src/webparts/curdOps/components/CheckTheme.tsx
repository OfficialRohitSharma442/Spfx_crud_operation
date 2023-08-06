import * as React from "react";
import { sp } from "@pnp/sp";
import ContextService from "../loc/ContextService";
const SharePointThemeName = () => {
  const [themeName, setThemeName] = React.useState("");

  React.useEffect(() => {
    // Initialize SharePoint PnP
    sp.setup({
      spfxContext: ContextService.GetFullContext(),
    });

    // Retrieve the theme name
    sp.web
      .select("ThemeInfo")
      .get()
      .then((web) => {
        const themeInfo = web.ThemeInfo;
        const themeName = themeInfo ? themeInfo.ThemeName : "";
        setThemeName(themeName);
      });
  }, []);

  return (
    <div>
      <h2>SharePoint Theme Name:</h2>
      <p>{themeName}</p>
    </div>
  );
};

function CheckTheme() {
  return <SharePointThemeName />;
}

export default CheckTheme;
