import * as React from 'react';
import { FormulaExplorer } from './FormulaExplorer';
import { ComparisonTool } from './ComparisonTool';
import { FormulaMapTool } from './FormulaMapTool';
import { DependentsTracer } from './DependentsTracer';
import { CalculationFlowTool } from './CalculationFlowTool';
import { Settings } from './Settings';
import { KeyboardShortcutManager } from '../../utils/keyboardShortcuts';

/* global Excel */

export interface AppProps {}

export interface AppState {
  activeTab: 'explorer' | 'comparison' | 'mapping' | 'dependents' | 'flow' | 'settings';
  isOfficeInitialized: boolean;
}

export class App extends React.Component<AppProps, AppState> {
  constructor(props: AppProps) {
    super(props);

    this.state = {
      activeTab: 'explorer',
      isOfficeInitialized: false
    };
  }

  componentDidMount() {
    // Initialize keyboard shortcuts
    KeyboardShortcutManager.initialize();

    // Register default shortcuts
    this.registerShortcuts();

    this.setState({ isOfficeInitialized: true });
  }

  registerShortcuts = () => {
    const config = KeyboardShortcutManager.getConfig();

    // Explore Formula (Ctrl+Q)
    KeyboardShortcutManager.registerShortcut(config.exploreFormula, () => {
      this.setState({ activeTab: 'explorer' });
    });

    // Navigate Back (Ctrl+Backspace)
    KeyboardShortcutManager.registerShortcut(config.navigateBack, async () => {
      await Excel.run(async (context) => {
        await KeyboardShortcutManager.navigateBack(context);
      });
    });

    // Set Reference (Ctrl+Shift+S)
    KeyboardShortcutManager.registerShortcut(config.setReference, () => {
      this.setState({ activeTab: 'comparison' });
    });

    // Toggle Formula Map (Ctrl+Shift+M)
    KeyboardShortcutManager.registerShortcut(config.toggleFormulaMap, () => {
      this.setState({ activeTab: 'mapping' });
    });

    // Trace Dependents (Ctrl+Shift+Q)
    KeyboardShortcutManager.registerShortcut(config.traceDependents, () => {
      this.setState({ activeTab: 'dependents' });
    });

    // Calculation Flow (Ctrl+Shift+F)
    KeyboardShortcutManager.registerShortcut(config.calculationFlow, () => {
      this.setState({ activeTab: 'flow' });
    });
  };

  setActiveTab = (tab: AppState['activeTab']) => {
    this.setState({ activeTab: tab });
  };

  render() {
    const { activeTab, isOfficeInitialized } = this.state;

    if (!isOfficeInitialized) {
      return (
        <div className="loading">
          <div className="spinner"></div>
          <p>Loading Numeriq...</p>
        </div>
      );
    }

    return (
      <div className="app-container">
        <header className="app-header">
          <h1 className="app-title">Numeriq</h1>
          <p className="app-subtitle">Formula Explorer & Analysis</p>
        </header>

        <nav className="app-nav">
          <button
            className={`nav-button ${activeTab === 'explorer' ? 'active' : ''}`}
            onClick={() => this.setActiveTab('explorer')}
            title="Explore Formula (Ctrl+Q)"
          >
            <span className="nav-icon">🔍</span>
            <span className="nav-label">Explorer</span>
          </button>
          <button
            className={`nav-button ${activeTab === 'comparison' ? 'active' : ''}`}
            onClick={() => this.setActiveTab('comparison')}
            title="Compare Spreadsheets (Ctrl+Shift+S/C)"
          >
            <span className="nav-icon">⚖️</span>
            <span className="nav-label">Compare</span>
          </button>
          <button
            className={`nav-button ${activeTab === 'mapping' ? 'active' : ''}`}
            onClick={() => this.setActiveTab('mapping')}
            title="Formula Map (Ctrl+Shift+M)"
          >
            <span className="nav-icon">🗺️</span>
            <span className="nav-label">Map</span>
          </button>
          <button
            className={`nav-button ${activeTab === 'dependents' ? 'active' : ''}`}
            onClick={() => this.setActiveTab('dependents')}
            title="Trace Dependents (Ctrl+Shift+Q)"
          >
            <span className="nav-icon">🔗</span>
            <span className="nav-label">Trace</span>
          </button>
          <button
            className={`nav-button ${activeTab === 'flow' ? 'active' : ''}`}
            onClick={() => this.setActiveTab('flow')}
            title="Calculation Flow (Ctrl+Shift+F)"
          >
            <span className="nav-icon">📊</span>
            <span className="nav-label">Flow</span>
          </button>
          <button
            className={`nav-button ${activeTab === 'settings' ? 'active' : ''}`}
            onClick={() => this.setActiveTab('settings')}
          >
            <span className="nav-icon">⚙️</span>
            <span className="nav-label">Settings</span>
          </button>
        </nav>

        <main className="app-content">
          {activeTab === 'explorer' && <FormulaExplorer />}
          {activeTab === 'comparison' && <ComparisonTool />}
          {activeTab === 'mapping' && <FormulaMapTool />}
          {activeTab === 'dependents' && <DependentsTracer />}
          {activeTab === 'flow' && <CalculationFlowTool />}
          {activeTab === 'settings' && <Settings />}
        </main>

        <footer className="app-footer">
          <p className="footer-text">
            Press <kbd>Ctrl+Q</kbd> to explore | <kbd>Ctrl+Backspace</kbd> to navigate back
          </p>
        </footer>
      </div>
    );
  }
}
