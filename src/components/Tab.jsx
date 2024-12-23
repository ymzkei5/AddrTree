import { useContext } from "react";
import { Main } from "./addrTree/Main.tsx";
import { TeamsFxContext } from "./Context";

export default function Tab() {
  const { themeString } = useContext(TeamsFxContext);
  return (
    <div
      className={themeString === "default" ? "light" : themeString === "dark" ? "dark" : "contrast"}
    >
      <Main />
    </div>
  );
}