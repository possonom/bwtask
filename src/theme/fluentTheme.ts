import { 
  createTheme, 
  ITheme, 
  IPalette, 
  ISemanticColors,
  DefaultEffects,
  IEffects
} from '@fluentui/react';

// Custom palette extending Fluent UI 8 theme system
export const customPalette: Partial<IPalette> = {
  themePrimary: '#0078d4',
  themeLighterAlt: '#eff6fc',
  themeLighter: '#deecf9',
  themeLight: '#c7e0f4',
  themeTertiary: '#71afe5',
  themeSecondary: '#2b88d8',
  themeDarkAlt: '#106ebe',
  themeDark: '#005a9e',
  themeDarker: '#004578',
  neutralLighterAlt: '#faf9f8',
  neutralLighter: '#f3f2f1',
  neutralLight: '#edebe9',
  neutralQuaternaryAlt: '#e1dfdd',
  neutralQuaternary: '#d0d0d0',
  neutralTertiaryAlt: '#c8c6c4',
  neutralTertiary: '#a19f9d',
  neutralSecondary: '#605e5c',
  neutralPrimaryAlt: '#3b3a39',
  neutralPrimary: '#323130',
  neutralDark: '#201f1e',
  black: '#000000',
  white: '#ffffff',
};

// Custom semantic colors for better UX
export const customSemanticColors: Partial<ISemanticColors> = {
  bodyBackground: customPalette.white,
  bodyText: customPalette.neutralPrimary,
  bodySubtext: customPalette.neutralSecondary,
  actionLink: customPalette.themePrimary,
  actionLinkHovered: customPalette.themeDarkAlt,
  link: customPalette.themePrimary,
  linkHovered: customPalette.themeDarkAlt,
  buttonBackground: customPalette.neutralLighter,
  buttonBackgroundHovered: customPalette.neutralLight,
  buttonBackgroundPressed: customPalette.neutralQuaternaryAlt,
  buttonText: customPalette.neutralPrimary,
  buttonTextHovered: customPalette.neutralDark,
  buttonTextPressed: customPalette.neutralDark,
  primaryButtonBackground: customPalette.themePrimary,
  primaryButtonBackgroundHovered: customPalette.themeDarkAlt,
  primaryButtonBackgroundPressed: customPalette.themeDark,
  primaryButtonText: customPalette.white,
  inputBackground: customPalette.white,
  inputBorder: customPalette.neutralQuaternary,
  inputBorderHovered: customPalette.neutralTertiary,
  inputText: customPalette.neutralPrimary,
  errorText: '#d13438',
  warningText: '#ff8c00',
  successText: '#107c10',
};

// Custom effects for modern look
export const customEffects: Partial<IEffects> = {
  ...DefaultEffects,
  elevation4: '0 1.6px 3.6px 0 rgba(0, 0, 0, 0.132), 0 0.3px 0.9px 0 rgba(0, 0, 0, 0.108)',
  elevation8: '0 3.2px 7.2px 0 rgba(0, 0, 0, 0.132), 0 0.6px 1.8px 0 rgba(0, 0, 0, 0.108)',
  elevation16: '0 6.4px 14.4px 0 rgba(0, 0, 0, 0.132), 0 1.2px 3.6px 0 rgba(0, 0, 0, 0.108)',
  roundedCorner2: '2px',
  roundedCorner4: '4px',
  roundedCorner6: '6px',
};

// Create the main theme
export const fluentTheme: ITheme = createTheme({
  palette: customPalette,
  semanticColors: customSemanticColors,
  effects: customEffects,
  fonts: {
    small: {
      fontSize: '12px',
      fontWeight: 400,
    },
    medium: {
      fontSize: '14px',
      fontWeight: 400,
    },
    mediumPlus: {
      fontSize: '16px',
      fontWeight: 400,
    },
    large: {
      fontSize: '18px',
      fontWeight: 400,
    },
    xLarge: {
      fontSize: '20px',
      fontWeight: 600,
    },
    xxLarge: {
      fontSize: '28px',
      fontWeight: 600,
    },
  },
});

// Theme variants for different contexts
export const darkTheme: ITheme = createTheme({
  palette: {
    ...customPalette,
    neutralLighterAlt: '#282828',
    neutralLighter: '#313131',
    neutralLight: '#3f3f3f',
    neutralQuaternaryAlt: '#484644',
    neutralQuaternary: '#605e5c',
    neutralTertiaryAlt: '#a19f9d',
    neutralTertiary: '#c8c6c4',
    neutralSecondary: '#d0d0d0',
    neutralPrimaryAlt: '#dadada',
    neutralPrimary: '#ffffff',
    neutralDark: '#f4f4f4',
    black: '#f8f8f8',
    white: '#1f1f1f',
  },
  semanticColors: {
    ...customSemanticColors,
    bodyBackground: '#1f1f1f',
    bodyText: '#ffffff',
    bodySubtext: '#d0d0d0',
  },
  effects: customEffects,
});

// Utility function to get theme based on SharePoint context
export const getContextualTheme = (isDark: boolean = false): ITheme => {
  return isDark ? darkTheme : fluentTheme;
};

// Theme tokens for consistent spacing and sizing
export const themeTokens = {
  spacing: {
    xs: '4px',
    s: '8px',
    m: '16px',
    l: '24px',
    xl: '32px',
    xxl: '40px',
  },
  borderRadius: {
    small: '2px',
    medium: '4px',
    large: '6px',
  },
  shadows: {
    depth4: customEffects.elevation4,
    depth8: customEffects.elevation8,
    depth16: customEffects.elevation16,
  },
};
