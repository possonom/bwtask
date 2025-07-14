import * as React from 'react';
import { useState, useCallback } from 'react';
import { IHelloWorldProps } from './IHelloWorldProps';
import { 
  CommandBar, 
  ICommandBarItemProps, 
  TextField, 
  ThemeProvider,
  Stack,
  IStackTokens,
  PrimaryButton,
  DefaultButton,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  Panel,
  PanelType,
  Toggle,
  Dropdown,
  IDropdownOption,
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  SelectionMode
} from '@fluentui/react';
import { ICommandBarSampleDropDownProps, CommandBarSampleDropDown } from './CommandBarSampleDropDown';

interface ISampleData {
  id: number;
  name: string;
  status: string;
  date: Date;
}

const HelloWorld: React.FC<IHelloWorldProps> = (props) => {
  const { theme, description } = props;

  const stackTokens: IStackTokens = { childrenGap: 20 };
  const [counter, setCounter] = useState<number>(1);
  const [userName, setUserName] = useState<string>('');
  const [showMessage, setShowMessage] = useState<boolean>(false);
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [isPanelOpen, setIsPanelOpen] = useState<boolean>(false);
  const [isDarkMode, setIsDarkMode] = useState<boolean>(false);
  const [selectedOption, setSelectedOption] = useState<string>('option1');

  // Sample data for the list
  const [sampleData] = useState<ISampleData[]>([
    { id: 1, name: 'Project Alpha', status: 'Active', date: new Date('2024-01-15') },
    { id: 2, name: 'Project Beta', status: 'Completed', date: new Date('2024-02-20') },
    { id: 3, name: 'Project Gamma', status: 'Pending', date: new Date('2024-03-10') },
    { id: 4, name: 'Project Delta', status: 'Active', date: new Date('2024-01-25') },
  ]);

  const onButtonClick = useCallback((): void => {
    setCounter(prevState => prevState + 1);
    setShowMessage(true);
    setTimeout(() => setShowMessage(false), 3000);
  }, []);

  const onAsyncAction = useCallback(async (): Promise<void> => {
    setIsLoading(true);
    // Simulate async operation
    await new Promise(resolve => setTimeout(resolve, 2000));
    setIsLoading(false);
    setShowMessage(true);
    setTimeout(() => setShowMessage(false), 3000);
  }, []);

  const onUserNameChange = useCallback((event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
    setUserName(newValue || '');
  }, []);

  const onToggleChange = useCallback((event: React.MouseEvent<HTMLElement>, checked?: boolean): void => {
    setIsDarkMode(checked || false);
  }, []);

  const dropdownOptions: IDropdownOption[] = [
    { key: 'option1', text: 'Option 1' },
    { key: 'option2', text: 'Option 2' },
    { key: 'option3', text: 'Option 3' },
  ];

  const onDropdownChange = useCallback((event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    if (option) {
      setSelectedOption(option.key as string);
    }
  }, []);

  // Command bar configuration
  const dropdownSample: ICommandBarSampleDropDownProps = {
    text: 'My View 1',
    key: 'view1',
    tooltipText: 'Switch between different views',
    commandBarButtonAs: CommandBarSampleDropDown,
  };

  const commandBarItems: ICommandBarItemProps[] = [
    {
      key: 'new',
      text: 'New',
      iconProps: { iconName: 'Add' },
      onClick: () => setIsPanelOpen(true),
    },
    {
      key: 'refresh',
      text: 'Refresh',
      iconProps: { iconName: 'Refresh' },
      onClick: onAsyncAction,
    },
  ];

  const farItems: ICommandBarItemProps[] = [dropdownSample];

  // DetailsList columns
  const columns: IColumn[] = [
    {
      key: 'name',
      name: 'Name',
      fieldName: 'name',
      minWidth: 150,
      maxWidth: 200,
      isResizable: true,
    },
    {
      key: 'status',
      name: 'Status',
      fieldName: 'status',
      minWidth: 100,
      maxWidth: 150,
      isResizable: true,
      onRender: (item: ISampleData) => (
        <span style={{ 
          color: item.status === 'Active' ? '#107c10' : 
                 item.status === 'Completed' ? '#0078d4' : '#d13438' 
        }}>
          {item.status}
        </span>
      ),
    },
    {
      key: 'date',
      name: 'Date',
      fieldName: 'date',
      minWidth: 120,
      maxWidth: 180,
      isResizable: true,
      onRender: (item: ISampleData) => item.date.toLocaleDateString(),
    },
  ];

  return (
    <ThemeProvider theme={theme}>
      <div style={{ padding: '20px', backgroundColor: isDarkMode ? '#1f1f1f' : '#ffffff' }}>
        <h2>Hello World - SPFx Sample Component</h2>
        <p>{description}</p>

        {/* Command Bar */}
        <CommandBar
          items={commandBarItems}
          farItems={farItems}
          ariaLabel="Sample command bar"
        />

        {/* Message Bar */}
        {showMessage && (
          <MessageBar
            messageBarType={MessageBarType.success}
            isMultiline={false}
            onDismiss={() => setShowMessage(false)}
            dismissButtonAriaLabel="Close"
          >
            Action completed successfully! Counter is now: {counter}
          </MessageBar>
        )}

        <Stack tokens={stackTokens}>
          {/* Basic Controls */}
          <Stack horizontal tokens={{ childrenGap: 20 }} verticalAlign="center">
            <div>Counter: <strong>{counter}</strong></div>
            <PrimaryButton 
              data-testid="hw-main-btn-default" 
              text="Increment" 
              onClick={onButtonClick}
              disabled={isLoading}
            />
            <DefaultButton 
              text="Async Action" 
              onClick={onAsyncAction}
              disabled={isLoading}
            />
            {isLoading && <Spinner size={SpinnerSize.medium} />}
          </Stack>

          {/* Form Controls */}
          <Stack horizontal tokens={{ childrenGap: 20 }}>
            <TextField 
              data-testid="hw-main-txt-name" 
              label="User Name" 
              value={userName}
              onChange={onUserNameChange}
              placeholder="Enter your name"
              styles={{ root: { width: 200 } }}
            />
            <Dropdown
              label="Select Option"
              options={dropdownOptions}
              selectedKey={selectedOption}
              onChange={onDropdownChange}
              styles={{ root: { width: 150 } }}
            />
            <Toggle
              label="Dark Mode"
              checked={isDarkMode}
              onChange={onToggleChange}
            />
          </Stack>

          {/* Greeting */}
          {userName && (
            <div style={{ 
              padding: '10px', 
              backgroundColor: isDarkMode ? '#333' : '#f3f2f1',
              borderRadius: '4px'
            }}>
              Hello, <strong>{userName}</strong>! Welcome to SPFx development.
            </div>
          )}

          {/* Data List */}
          <div>
            <h3>Sample Data List</h3>
            <DetailsList
              items={sampleData}
              columns={columns}
              layoutMode={DetailsListLayoutMode.justified}
              selectionMode={SelectionMode.none}
              isHeaderVisible={true}
            />
          </div>
        </Stack>

        {/* Panel */}
        <Panel
          headerText="New Item Panel"
          isOpen={isPanelOpen}
          onDismiss={() => setIsPanelOpen(false)}
          type={PanelType.medium}
          closeButtonAriaLabel="Close"
        >
          <Stack tokens={stackTokens}>
            <TextField label="Item Name" placeholder="Enter item name" />
            <Dropdown
              label="Status"
              options={[
                { key: 'active', text: 'Active' },
                { key: 'pending', text: 'Pending' },
                { key: 'completed', text: 'Completed' },
              ]}
            />
            <Stack horizontal tokens={{ childrenGap: 10 }}>
              <PrimaryButton text="Save" onClick={() => setIsPanelOpen(false)} />
              <DefaultButton text="Cancel" onClick={() => setIsPanelOpen(false)} />
            </Stack>
          </Stack>
        </Panel>
      </div>
    </ThemeProvider>
  );
};

export default HelloWorld;
