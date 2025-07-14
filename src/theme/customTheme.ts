import { baseTheme } from "./basePalette";
import { DefaultEffects, IPalette, ISemanticColors, ITheme, PartialTheme, createTheme } from "@fluentui/react/lib/Theme";
import { getComponentStyleOverrides, getEffectStyleOverrides } from "./componentOverrides";
/**
 * We load our theme based on basic styles and current styles from window
 * how to use styles: https://github.com/microsoft/fluentui/blob/master/packages/merge-styles/README.md
 * @returns ITheme
 */

export const loadCustomTheme = (): ITheme => {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const themeColorsFromWindow: Partial<IPalette> = (window as any).__themeState__.theme;

  // combine base theme and theme from sharepoint page
  const fullPalette = { ...baseTheme, ...themeColorsFromWindow };

  const effects = getEffectStyleOverrides();
  const fullEffects = { ...DefaultEffects, ...effects };
  const componentStyles = getComponentStyleOverrides(fullPalette, fullEffects);

  const semanticColors: Partial<ISemanticColors> = {};

  const fullTheme: PartialTheme = {
    palette: fullPalette,
    effects: fullEffects,
    fonts: {
      // this seems a bug, why the property not here but used in components
      // eslint-disable-next-line @typescript-eslint/ban-ts-comment
      //@ts-ignore
      xLargePlus: "14px",
    },
    semanticColors: semanticColors,
    components: componentStyles,
  };

  return createTheme(fullTheme);
};
