# SPFx Template Project

A comprehensive SharePoint Framework (SPFx) template with sample code, utilities, and best practices for modern SharePoint development.

## 🚀 Features

### Web Parts
- **HelloWorld**: Comprehensive sample component demonstrating Fluent UI controls, state management, and user interactions
- **SpDataAccess**: Advanced data access patterns using PnP V2 with custom hooks, filtering, and CRUD operations

### Utilities & Services
- **Error Handling**: Robust error handling with user-friendly messages and detailed logging
- **SharePoint Permissions**: Helper functions for checking user permissions and access levels
- **Project Service**: Full CRUD operations for SharePoint list management
- **Custom Hooks**: Reusable hooks for data fetching and state management

### UI Components
- **Command Bar**: Custom dropdown components with contextual menus
- **Data Lists**: Advanced DetailsList with filtering, sorting, and selection
- **Forms**: Comprehensive form controls with validation and feedback
- **Panels**: Side panels for detailed views and editing
- **Responsive Design**: Mobile-friendly layouts and interactions

## 🛠 Tech Stack

- **SharePoint Framework (SPFx)**: 1.4.0
- **React**: 17.0.2
- **TypeScript**: 4.9.5
- **Fluent UI**: 8.110.4
- **PnP JS**: 2.0.9
- **Jest**: Testing framework
- **ESLint**: Code linting

## 📦 Installation

```bash
# Install dependencies
npm install

# Start development server (German locale)
npm run serve:de

# Start development server (US locale)
npm run serve:us

# Build for development
npm run build:dev

# Build for production
npm run build:prod

# Package solution
npm run package

# Run tests
npm test
```

## 🏗 Project Structure

```
src/
├── webparts/
│   ├── helloWorld/           # Sample web part with UI controls
│   └── spDataAccess/         # Data access demonstration
├── services/
│   ├── useQueryProjects.ts   # Custom hook for data fetching
│   └── projectService.ts     # CRUD operations service
├── helper/
│   ├── errorHelper.ts        # Error handling utilities
│   └── spPermissions.ts      # SharePoint permissions helper
├── interfaces/
│   ├── Project.ts            # Project data models
│   └── Status.ts             # Status enums and utilities
└── __tests__/                # Test files
```

## 🎯 Key Components

### HelloWorld Web Part
Demonstrates:
- Fluent UI controls (buttons, text fields, dropdowns, toggles)
- State management with React hooks
- Command bar with custom dropdowns
- Data lists with DetailsList
- Panels for detailed views
- Loading states and async operations
- Dark mode toggle
- Message notifications

### SpDataAccess Web Part
Features:
- SharePoint list data fetching with PnP V2
- Custom hooks for data management
- Advanced filtering and search
- CRUD operations
- Error handling with fallback to mock data
- Permission checking
- Responsive data display

### Utility Services

#### Error Helper
- Comprehensive error message extraction
- SharePoint-specific error handling
- User-friendly error formatting
- Detailed logging with context

#### SharePoint Permissions
- User permission checking
- List access validation
- CRUD permission verification
- Site-level permission queries

#### Project Service
- Full CRUD operations
- Batch operations
- Advanced filtering
- Summary statistics
- Type-safe interfaces

## 🧪 Testing

The project includes comprehensive test coverage:

```bash
# Run all tests
npm test

# Run tests in watch mode
npm test -- --watch

# Run tests with coverage
npm test -- --coverage
```

Test files are located alongside their source files with `.test.ts` or `.test.tsx` extensions.

## 🔧 Configuration

### Environment Setup
1. Ensure Node.js 14+ is installed
2. Install SharePoint Framework globally: `npm install -g @microsoft/generator-sharepoint`
3. Clone and install dependencies
4. Configure SharePoint connection in `config/serve.json`

### SharePoint List Setup
The SpDataAccess web part expects a "Projects" list with the following columns:
- Title (Single line of text)
- Status (Choice: Not Started, In Progress, Completed, On Hold)
- Description (Multiple lines of text)
- StartDate (Date and Time)
- EndDate (Date and Time)
- Owner (Single line of text)
- Priority (Number)
- Progress (Number)
- Budget (Number)
- Tags (Single line of text)

## 🎨 Styling

The project uses:
- Fluent UI theme system
- SCSS for custom styling
- Responsive design patterns
- Dark mode support
- Consistent spacing and typography

## 📝 Best Practices

### Code Organization
- Modular component structure
- Custom hooks for reusable logic
- Type-safe interfaces
- Comprehensive error handling
- Consistent naming conventions

### Performance
- Lazy loading for large datasets
- Memoization for expensive operations
- Efficient re-rendering patterns
- Optimized bundle size

### Accessibility
- ARIA labels and roles
- Keyboard navigation support
- Screen reader compatibility
- High contrast support

## 🚀 Deployment

```bash
# Build and package for production
npm run release

# Deploy to SharePoint app catalog
# Upload the .sppkg file from sharepoint/solution/
```

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests for new functionality
5. Ensure all tests pass
6. Submit a pull request

## 📄 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 🆘 Support

For issues and questions:
1. Check the existing issues in the repository
2. Review the SharePoint Framework documentation
3. Consult the Fluent UI documentation
4. Create a new issue with detailed information

## 🔗 Useful Links

- [SharePoint Framework Documentation](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/)
- [Fluent UI Documentation](https://developer.microsoft.com/en-us/fluentui)
- [PnP JS Documentation](https://pnp.github.io/pnpjs/)
- [React Documentation](https://reactjs.org/docs/)
- [TypeScript Documentation](https://www.typescriptlang.org/docs/)
