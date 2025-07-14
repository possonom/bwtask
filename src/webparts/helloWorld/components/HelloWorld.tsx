import * as React from 'react';
  import { useState } from 'react';
  import { IHelloWorldProps } from './IHelloWorldProps';
  import { CommandBar, ICommandBarItemProps, TextField, ThemeProvider, } from '@fluentui/react';
  import { Stack, IStackTokens } from '@fluentui/react';
  import { PrimaryButton } from '@fluentui/react/lib/Button';
import { ICommandBarSampleDropDownProps, CommandBarSampleDropDown } from './CommandBarSampleDropDown';



  const HelloWorld: React.FC<IHelloWorldProps> = (props) => {
    const {
      theme
    } = props;

    const stackTokens: IStackTokens = { childrenGap: 40 };
    const [counter, setCounter] = useState<number>(1);

    const onButtonClick = (): void => {
      setCounter(prevState => prevState + 1);
    };
    const dropdownSample: ICommandBarSampleDropDownProps = {
      text: 'Meine Ansicht 1',
      key: 'test',
      tooltipText: 'Tooltip',
      commandBarButtonAs: CommandBarSampleDropDown,
    };
    const farItems: ICommandBarItemProps[] = [dropdownSample];
    
    return (
      <ThemeProvider theme={theme}>
         <CommandBar
            items={[]}
            farItems={farItems}
            ariaLabel={"TODO"}
        />
        <br></br>
        <Stack horizontal tokens={stackTokens}>
       
            <div>Counter: {counter}</div>
            <PrimaryButton data-testid="hw-main-btn-default" text="Increment" onClick={onButtonClick} />
        </Stack>
        <TextField data-testid="hw-main-txt-name" label='Test' />
      </ThemeProvider>
    );
  };

  export default HelloWorld;
