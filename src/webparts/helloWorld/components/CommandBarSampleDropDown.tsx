import { CommandBarButton } from "@fluentui/react/lib/components/Button/CommandBarButton/CommandBarButton";
import { ICommandBarItemProps } from "@fluentui/react/lib/components/CommandBar";
import { useId } from "@fluentui/react-hooks";
import { IButtonStyles } from "@fluentui/react/lib/components/Button";
import {
  ITooltipHostStyles,
  TooltipHost,
} from "@fluentui/react/lib/components/Tooltip";
import * as React from "react";
import { mergeStyleSets } from "@fluentui/react/lib/Styling";
import {
  ContextualMenuItemType,
  IContextualMenuItem,
} from "@fluentui/react/lib/components/ContextualMenu";

export interface ICommandBarSampleDropDownProps extends ICommandBarItemProps {
  text: string;
  tooltipText: string;
}

const saveViewCommandStyle: Partial<IButtonStyles> = {
  label: {
    maxWidth: 120,
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
    overflow: "hidden",
  },
};
const hostStyles: Partial<ITooltipHostStyles> = { root: { display: "flex" } };
export const contextualMenuStyles = mergeStyleSets({
  headerText: { fontWeight: 600, FontSize: 14 },
  headerIcon: { display: "none" },
});

/**
 * Component to render a view selector as sample
 * @param commandProps
 * @returns
 */
export const CommandBarSampleDropDown = (commandProps: ICommandBarSampleDropDownProps) => {
  const tooltipId = useId("tooltipId");

  const { text, tooltipText, key } = commandProps;
  const viewsItems: IContextualMenuItem[] = [
   
    {
      key: "samplekey",
      text: text,
      itemType: ContextualMenuItemType.Normal,
      iconProps: {
        styles: {
          root: [
            contextualMenuStyles.headerIcon,
            contextualMenuStyles.headerText,
          ],
        },
      },
    },
  ];
  return (
    <TooltipHost
      key={key}
      content={tooltipText}
      id={tooltipId}
      styles={hostStyles}
    >
      <CommandBarButton
        data-aria-describedby={tooltipId}
        text={text}
        iconProps={{ iconName: "AlignLeft" }}
        menuProps={{
          styles: { container: { minWidth: 250 }, header: { fontSize: 14 } },
          items: viewsItems,
        }}
        styles={saveViewCommandStyle}
      />
    </TooltipHost>
  );
};
