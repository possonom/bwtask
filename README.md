# SPFx Template
This is our template project for SPFx solutions. Please use it to create your SPFx solutions.

The default SPFx solution generated with yeoman sp generator 1.10.0 for SP2019 was adapted as follows:

- TypeScript 4.9.5
- React 17.0.2
- pnp 2.0.9
- FluentUI 8.x
- tslint removed, replaced by typescript-eslint with additional rules for react

## Update your existing solution

Instructions are detailled here: [Manually update existing solutions](./manually-update-existing-solutions.md)

## NPM scripts

The following npm scripts are available per default:

| script        | functionality | variants |
|---------------|---------------|----------|
| `clean` | remove the folders "lib", "temp", "dist" and <br>"sharepoint" | - |
| `test` | run the jest unit tests | - |
| `serve` | build the solution and start a server for the <br>SP-workbench.<br>- browser doesn't open automatically<br>- default locale is "de-de"  | `serve:de`<br>`serve:en`   |
| `build` | *bundle* the solution (shipment/prod) | `build:dev`<br>`build:prod` |
| `package` | run the *package-solution* gulp task (shipment/prod) | - |
| `release` | build, bundle and package the solution<br>for release/shipment<br>i.e. release the SPFx package for (production) deployment | - |

run the script per:

```shell
npm run <NAME-OF-THE-SCRIPT>
```

## Styling
### FluentUI
FluentUI needs a custom theme, this is defined in /constants and used with the ThemeProvider
### Custom SCSS Modules
In /css, you'll find the default BW Colors. Import them in your SCSS file using @use, not @import:

    @use '../../../css/BwColors.module.scss';

## Logging
Please use the included PnP Logger and make sure to adapt the activeLogLevel to the environment where the solution is deployed, i.e. use LogLevel.Error for PROD

## Linting
Please see .eslintrc.js for configuration & rules
Please make sure that the same rules are enforced in every solution for multi-solution projects!

## Testing (Planned - not yet available)
We are using Jest as test runner and React Testing Library to test React components

Philosophy of the testing library:

    The more your tests resemble the way your software is used, the more confidence they can give you.

https://testing-library.com/docs/


## Adding more WebParts

In case you add more webparts to your SPFx solution using yeoman, the react version in the package.json will be reset to 15.x by yeoman. You have to re-install the 17.0.2 version using npm.

    npm install react@17.0.2 react-dom@17.0.2
    npm install @types/react@17.0.2 @types/react-dom@17.0.2 --save-dev

In addition, the react types must be changed again:

"resolutions": {
    "@types/react": "17.0.2"
}

## pnp Version 2.0.9
### Packages
Common: https://pnp.github.io/pnpjs/v2/common/
Logging: https://pnp.github.io/pnpjs/v2/logging/
OData: https://pnp.github.io/pnpjs/v2/odata/

## Using template
To use the template, the following values must be changed:
package-solution.json -> solution -> id
package-solution.json -> solution -> name

Maybe also the zippedPackage
package-solution.json -> paths -> zippedPackage

SpDataAccessWebPart.manifest.json -> id 
HelloWorldWebPart.manifest.json -> id

