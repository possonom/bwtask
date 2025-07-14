import { IStyleFunctionOrObject } from '@fluentui/merge-styles';
import { ComponentsStyles, IEffects, IPalette } from '@fluentui/react/lib/Theme';
import { IButtonStyles } from '@fluentui/react/lib/components/Button';
import { ICommandBarStyles } from '@fluentui/react/lib/components/CommandBar';
import { IDetailsColumnStyleProps, IDetailsColumnStyles, IDetailsListStyleProps, IDetailsListStyles } from '@fluentui/react/lib/components/DetailsList';
import { ILinkStyles } from '@fluentui/react/lib/components/Link';
import { IPivotStyleProps, IPivotStyles } from '@fluentui/react/lib/components/Pivot';

/**
 * Override global Effects like shadows
 * https://github.com/microsoft/fluentui/tree/master/packages/react/src/components
 * @returns
 */
export const getEffectStyleOverrides = (): Partial<IEffects> => {
    const boxShadowOverride = '0px 4px 8px 0px rgba(0, 0, 0, 0.3)';
    const effects: Partial<IEffects> = {
        elevation8: boxShadowOverride, //Tooltip or Dropdown callout
        elevation16: boxShadowOverride, //Calloutoverride
    };
    return effects;
};

/**
 * Overrides styles for components, you can have a better look on components styling and what to override on:
 * https://github.com/microsoft/fluentui/tree/master/packages/react/src/components
 * @param themePalette
 * @returns
 */
export const getComponentStyleOverrides = (themePalette: Partial<IPalette>, effects: Partial<IEffects>): ComponentsStyles => {
    const componentstyles: ComponentsStyles = {
        Link: {
            styles: {
                root: {
                    textDecoration: 'none',
                    color: themePalette.themePrimary,
                    selectors: {
                        '&:hover': {
                            color: themePalette.themeDarker,
                        },
                    },
                },
            } as ILinkStyles,
        },

        CommandBar: {
            styles: { root: { background: themePalette.neutralLighter, padding: '0px 14px 0px 14px' } } as ICommandBarStyles,
        },
        CommandBarButton: {
            styles: {
                root: {
                    backgroundColor: 'transparent',
                },
                rootHovered: {
                    color: themePalette.neutralDark,
                    backgroundColor: themePalette.neutralLight,
                },
            } as IButtonStyles,
        },

        DetailsList: {
            styles: {
                root: { fontSize: 12, overflowX: 'hidden', overflowY: 'auto', maxHeight: 490 },
            } as IStyleFunctionOrObject<IDetailsListStyleProps, IDetailsListStyles>,
        },
        DetailsHeader: {
            styles: {
                root: {
                    borderBottomColor: themePalette.neutralLight,
                },
            },
        },
        DetailsColumn: {
            styles: { root: { overflow: 'hidden' }, cellTitle: { fontSize: 12 }, cellName: { fontSize: 12, fontWeight: 400 } } as IStyleFunctionOrObject<
                IDetailsColumnStyleProps,
                IDetailsColumnStyles
            >,
        },
        Pivot: {
            styles: {
                link: { borderRadius: effects.roundedCorner2 },
                linkIsSelected: { borderRadius: effects.roundedCorner2 },
            } as IStyleFunctionOrObject<IPivotStyleProps, IPivotStyles>,
        },
    };

    return componentstyles;
};
