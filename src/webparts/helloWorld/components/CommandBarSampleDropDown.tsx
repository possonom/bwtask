import * as React from 'react';
import { useState, useCallback } from 'react';
import {
  IContextualMenuProps,
  IContextualMenuItem,
  DirectionalHint,
  ContextualMenu,
  ICommandBarItemProps,
  DefaultButton,
  Icon,
  Stack,
  Text,
} from '@fluentui/react';
import { fluentTheme } from '../../../theme/fluentTheme';

export interface ICommandBarSampleDropDownProps extends ICommandBarItemProps {
  text: string;
  key: string;
  tooltipText?: string;
}

export const CommandBarSampleDropDown: React.FC<ICommandBarSampleDropDownProps> = (props) => {
  const [isMenuVisible, setIsMenuVisible] = useState<boolean>(false);
  const [target, setTarget] = useState<HTMLElement | null>(null);

  const onButtonClick = useCallback((event: React.MouseEvent<HTMLElement>) => {
    setTarget(event.currentTarget as HTMLElement);
    setIsMenuVisible(true);
  }, []);

  const onMenuDismiss = useCallback(() => {
    setIsMenuVisible(false);
    setTarget(null);
  }, []);

  const menuItems: IContextualMenuItem[] = [
    {
      key: 'view1',
      text: 'My View 1',
      iconProps: { iconName: 'View' },
      onClick: () => {
        console.log('View 1 selected');
        onMenuDismiss();
      },
    },
    {
      key: 'view2',
      text: 'My View 2',
      iconProps: { iconName: 'ViewList' },
      onClick: () => {
        console.log('View 2 selected');
        onMenuDismiss();
      },
    },
    {
      key: 'view3',
      text: 'My View 3',
      iconProps: { iconName: 'ViewDashboard' },
      onClick: () => {
        console.log('View 3 selected');
        onMenuDismiss();
      },
    },
    {
      key: 'divider1',
      itemType: 1, // ContextualMenuItemType.Divider
    },
    {
      key: 'settings',
      text: 'View Settings',
      iconProps: { iconName: 'Settings' },
      onClick: () => {
        console.log('Settings selected');
        onMenuDismiss();
      },
    },
  ];

  const contextualMenuProps: IContextualMenuProps = {
    items: menuItems,
    target: target,
    onDismiss: onMenuDismiss,
    directionalHint: DirectionalHint.bottomLeftEdge,
    styles: {
      root: {
        backgroundColor: fluentTheme.semanticColors.bodyBackground,
        border: `1px solid ${fluentTheme.semanticColors.inputBorder}`,
        borderRadius: '4px',
        boxShadow: '0 4px 8px rgba(0, 0, 0, 0.1)',
      },
    },
  };

  return (
    <>
      <DefaultButton
        onClick={onButtonClick}
        title={props.tooltipText}
        styles={{
          root: {
            backgroundColor: 'transparent',
            border: 'none',
            minWidth: 'auto',
            padding: '8px 12px',
          },
          rootHovered: {
            backgroundColor: fluentTheme.semanticColors.buttonBackgroundHovered,
          },
          rootPressed: {
            backgroundColor: fluentTheme.semanticColors.buttonBackgroundPressed,
          },
        }}
      >
        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 4 }}>
          <Text variant="medium">{props.text}</Text>
          <Icon iconName="ChevronDown" styles={{ root: { fontSize: '12px' } }} />
        </Stack>
      </DefaultButton>
      {isMenuVisible && <ContextualMenu {...contextualMenuProps} />}
    </>
  );
};
