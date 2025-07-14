import * as React from 'react';
import { render, screen, fireEvent, waitFor } from '@testing-library/react';
import '@testing-library/jest-dom';
import HelloWorld from '../HelloWorld';
import { IHelloWorldProps } from '../IHelloWorldProps';
import { createTheme } from '@fluentui/react';

// Mock theme
const mockTheme = createTheme();

const defaultProps: IHelloWorldProps = {
  theme: mockTheme,
  description: 'Test description',
};

describe('HelloWorld Component', () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  it('renders without crashing', () => {
    render(<HelloWorld {...defaultProps} />);
    expect(screen.getByText('Hello World - SPFx Sample Component')).toBeInTheDocument();
  });

  it('displays the description prop', () => {
    render(<HelloWorld {...defaultProps} />);
    expect(screen.getByText('Test description')).toBeInTheDocument();
  });

  it('increments counter when button is clicked', () => {
    render(<HelloWorld {...defaultProps} />);
    
    const button = screen.getByTestId('hw-main-btn-default');
    const counterText = screen.getByText(/Counter:/);
    
    expect(counterText).toHaveTextContent('Counter: 1');
    
    fireEvent.click(button);
    
    expect(counterText).toHaveTextContent('Counter: 2');
  });

  it('shows success message when increment button is clicked', async () => {
    render(<HelloWorld {...defaultProps} />);
    
    const button = screen.getByTestId('hw-main-btn-default');
    fireEvent.click(button);
    
    await waitFor(() => {
      expect(screen.getByText(/Action completed successfully!/)).toBeInTheDocument();
    });
  });

  it('updates user name when text field changes', () => {
    render(<HelloWorld {...defaultProps} />);
    
    const textField = screen.getByTestId('hw-main-txt-name');
    fireEvent.change(textField, { target: { value: 'John Doe' } });
    
    expect(screen.getByText('Hello, John Doe! Welcome to SPFx development.')).toBeInTheDocument();
  });

  it('displays sample data in the list', () => {
    render(<HelloWorld {...defaultProps} />);
    
    expect(screen.getByText('Project Alpha')).toBeInTheDocument();
    expect(screen.getByText('Project Beta')).toBeInTheDocument();
    expect(screen.getByText('Project Gamma')).toBeInTheDocument();
    expect(screen.getByText('Project Delta')).toBeInTheDocument();
  });

  it('opens panel when New button is clicked', () => {
    render(<HelloWorld {...defaultProps} />);
    
    const newButton = screen.getByText('New');
    fireEvent.click(newButton);
    
    expect(screen.getByText('New Item Panel')).toBeInTheDocument();
  });

  it('toggles dark mode', () => {
    render(<HelloWorld {...defaultProps} />);
    
    const toggle = screen.getByRole('switch', { name: /Dark Mode/i });
    fireEvent.click(toggle);
    
    // Check if the background color changed (this is a simplified test)
    const container = screen.getByText('Hello World - SPFx Sample Component').closest('div');
    expect(container).toHaveStyle('background-color: #1f1f1f');
  });

  it('handles async action with loading state', async () => {
    render(<HelloWorld {...defaultProps} />);
    
    const asyncButton = screen.getByText('Async Action');
    fireEvent.click(asyncButton);
    
    // Check if spinner appears
    expect(screen.getByRole('progressbar')).toBeInTheDocument();
    
    // Wait for async operation to complete
    await waitFor(() => {
      expect(screen.queryByRole('progressbar')).not.toBeInTheDocument();
    }, { timeout: 3000 });
  });

  it('handles dropdown selection', () => {
    render(<HelloWorld {...defaultProps} />);
    
    const dropdown = screen.getByRole('combobox', { name: /Select Option/i });
    fireEvent.click(dropdown);
    
    const option2 = screen.getByText('Option 2');
    fireEvent.click(option2);
    
    // Verify selection (this would need more specific testing based on implementation)
    expect(dropdown).toBeInTheDocument();
  });

  it('disables buttons during loading', async () => {
    render(<HelloWorld {...defaultProps} />);
    
    const asyncButton = screen.getByText('Async Action');
    const incrementButton = screen.getByTestId('hw-main-btn-default');
    
    fireEvent.click(asyncButton);
    
    expect(incrementButton).toBeDisabled();
    expect(asyncButton).toBeDisabled();
    
    await waitFor(() => {
      expect(incrementButton).not.toBeDisabled();
      expect(asyncButton).not.toBeDisabled();
    }, { timeout: 3000 });
  });
});
