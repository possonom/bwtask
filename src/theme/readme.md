# Styling and Theme with Fluent UI 8

Diese Biblothek hilft bei einem zentralisierten und verbesserten Styling Prozess von Fluent UI Komponenten

## Benutzung

Es wird im Grunde nur ein Themeprovider benötigt, dieser zieht sich das Design aus der basePalette.ts und aus dem Design vom SharePoint.

XXXXWebPart.ts

```ts
private theme: ITheme;
public async onInit(): Promise<void> {
   ...
   this.theme = loadCustomTheme();
}
...

public render(): void {
    const webPartProps: IWebPartProps = {
      ...
      theme: this.theme,
    };
    ...
    const element: React.ReactElement<IWebPartProps> = React.createElement(MainComponent, webPartProps);
    ReactDom.render(element, this.domElement);
}

```

MainComponent.ts

```ts
const MainComponent: React.FC<IWebPartProps> = (props) => {
  const { theme } = props;
  return (
    <ThemeProvider theme={theme}>
    ... Alle Komponenten innerhalb haben nun das Theme korrekt angewendet!
     </ThemeProvider>
  );
}
```

## Motvation
Wir benutzen als Basis einer Implementierung ein UI Design in Figma. Dieses hat leichte Anpassungen, welche mit über das Styling abgebildet werden sollen.

Dies wird ab Fluent UI 7 hauptsächlich über styleable components erreicht, diese haben eine styles prop, welche eine Anpassung des Designs ermöglichen. Eine genaue Erklärung findet Ihr hier:

https://github.com/microsoft/fluentui/blob/master/docs/react-wiki-archive/Component-Styling.md

Beschrieben ist auch, warum scss nicht mehr zur Anwendung findet:

**Motivations for moving away from SCSS**

>SCSS a build time process of expanding a high level css-like language into raw css. Our pipeline to load the raw css goes through a javascript conversion process and gets loaded on the page via a javascript library called load-themed-styles. Effectively, we have a complex build process which takes rules, converts them into JavaScript, and loads them dynamically.
>
>This process is complicated and adds a number of limitations.
>
>***We can't register classes dynamically***
>
>Scenarios like "make this area of the screen use a different theme" become really complicated if build time is the only time for evaluations.
>
>***Bundle size and css loading heft with scss***
>
>If a button has 20 different possible states, using scss you must load the css for all 20 of the states preemptively, so you end up loading way more rules than you will ever actually use. There is no "plt1 styles vs delay loaded styles." The best you can do is to partition your css to specific modules, and delay load the modules. But in this model, you will still preempt loading a lot of rules that aren't used.
>
>Sass also encourages "mixins" as a way to have one definition of styles that can be used in multiple places. This completely fights against bundle size, since mixins simply stamp duplicates copies of the same rules wherever they're used, resulting in bloated (but highly compressible) style definitions. The compression helps but all of this could be avoided by using a different approach to defining our styling.
>
>***Constant battle with specificity***
>
>Perhaps the most difficult thing to resolve is css specificity. Countless hacks have been implemented to "slightly tweak" styling of a thing in a particular context. If your rule is equally specific as an existing rule, you have a race condition; last one to register wins, resulting in hacks that only work sometimes. And even if your rule is more specific than an existing rule, there are no gates that can catch an existing rule being changed to be more specific later, resulting in breaking the workarounds.
>
>We want a system which allows users to pass in their overrides, which can create new permutations of classes which are only 1 level of specificity deep, providing a consistent safe way to override the defaults.

## Wie ergänze ich neue Anpassungen

Diese werden in der ComponentOverrides definiert

componentOverrides.ts

```ts
 const componentstyles: ComponentsStyles = {
    Link: {
      styles: {
        root: {
          textDecoration: "none",
          color: themePalette.themePrimary,
          ":hover": {
            color: themePalette.themeDarker,
          },
        },
      } as ILinkStyles,
    },
    ...
```
Jede Komponente kann hier eigens überschrieben werden. Was hier hier am besten überschreiben könnt, findet ihr in der original Implementierung in der styles Datei.

https://github.com/microsoft/fluentui/tree/master/packages/react/src/components

## Was sind Effekte?
Fluent Ui definiert einige Basiswerte über Effekte. Diese können so einfach überschrieben werden.

componentOverrides.ts

```ts
export const getEffectStyleOverrides = (): Partial<IEffects> => {
  const boxShadowOverride = "0px 4px 8px 0px rgba(0, 0, 0, 0.3)";
  const effects: Partial<IEffects> = {
    roundedCorner2: "0px",
    elevation8: boxShadowOverride, //Tooltip or Dropdown callout
    elevation16: boxShadowOverride, //Calloutoverride
  };
  ...
```

## Was ist die bwColorPalette.ts?

Diese definiert Basisfarben der BW

## Wie kann ich eigene Komponenten mit einem Themewert über Klassen stylen?

Die Methode mergeStyleSets hilft bei der Erzeugung von Styles oder Klassen:

https://github.com/microsoft/fluentui/blob/master/packages/merge-styles/README.md

```ts
const getWebPartTitleStyles = (theme: ITheme) => {
  return mergeStyleSets({
    root: {
      fontSize: 24,
      "> div": {
        fontSize: 24,
        color: theme.palette.neutralPrimary,
        fontWeight: 100,
      },
    },
  });
};

const Sample: React.FC<Props> = (props) => {
  //Achtung, das Theme sollte nicht immer wieder als Property in jeder komponente übergeben werden, bitte Hooks, Context oder einen Store benutzen!
  const {  theme } = props;
  const webTitleStyles = getWebPartTitleStyles(theme);
    return (
      <WebPartTitle ... className={webTitleStyles.root}  />
    )
}

```

