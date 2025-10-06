import * as React from 'react';
import { KeyboardShortcutManager, ShortcutConfig } from '../../utils/keyboardShortcuts';

export interface SettingsState {
  shortcuts: ShortcutConfig;
  explorerOptions: {
    closeOnEscape: boolean;
    autoNavigateBack: boolean;
  };
}

export class Settings extends React.Component<{}, SettingsState> {
  constructor(props: {}) {
    super(props);

    this.state = {
      shortcuts: KeyboardShortcutManager.getConfig(),
      explorerOptions: {
        closeOnEscape: true,
        autoNavigateBack: true
      }
    };
  }

  updateShortcut = (key: keyof ShortcutConfig, value: string) => {
    this.setState(
      prevState => ({
        shortcuts: {
          ...prevState.shortcuts,
          [key]: value
        }
      }),
      () => {
        KeyboardShortcutManager.updateConfig(this.state.shortcuts);
      }
    );
  };

  resetShortcuts = () => {
    const defaultConfig: ShortcutConfig = {
      exploreFormula: 'Ctrl+Q',
      navigateBack: 'Ctrl+Backspace',
      setReference: 'Ctrl+Shift+S',
      setComparator: 'Ctrl+Shift+C',
      toggleFormulaMap: 'Ctrl+Shift+M',
      traceDependents: 'Ctrl+Shift+Q',
      tracePrecedents: 'Ctrl+Q',
      calculationFlow: 'Ctrl+Shift+F'
    };

    this.setState({ shortcuts: defaultConfig }, () => {
      KeyboardShortcutManager.updateConfig(defaultConfig);
    });
  };

  clearNavigationHistory = () => {
    KeyboardShortcutManager.clearHistory();
    alert('Navigation history cleared');
  };

  render() {
    const { shortcuts, explorerOptions } = this.state;

    return (
      <div className="settings">
        <div className="settings-header">
          <h2>Settings</h2>
        </div>

        <div className="settings-section">
          <h3>Keyboard Shortcuts</h3>
          <p className="section-description">
            Customize keyboard shortcuts for Numeriq features
          </p>

          <div className="shortcut-list">
            <div className="shortcut-item">
              <label>
                <span className="shortcut-label">Explore Formula</span>
                <input
                  type="text"
                  value={shortcuts.exploreFormula}
                  onChange={(e) => this.updateShortcut('exploreFormula', e.target.value)}
                  className="shortcut-input"
                />
              </label>
            </div>

            <div className="shortcut-item">
              <label>
                <span className="shortcut-label">Navigate Back</span>
                <input
                  type="text"
                  value={shortcuts.navigateBack}
                  onChange={(e) => this.updateShortcut('navigateBack', e.target.value)}
                  className="shortcut-input"
                />
              </label>
            </div>

            <div className="shortcut-item">
              <label>
                <span className="shortcut-label">Set as Reference</span>
                <input
                  type="text"
                  value={shortcuts.setReference}
                  onChange={(e) => this.updateShortcut('setReference', e.target.value)}
                  className="shortcut-input"
                />
              </label>
            </div>

            <div className="shortcut-item">
              <label>
                <span className="shortcut-label">Set as Comparator</span>
                <input
                  type="text"
                  value={shortcuts.setComparator}
                  onChange={(e) => this.updateShortcut('setComparator', e.target.value)}
                  className="shortcut-input"
                />
              </label>
            </div>

            <div className="shortcut-item">
              <label>
                <span className="shortcut-label">Toggle Formula Map</span>
                <input
                  type="text"
                  value={shortcuts.toggleFormulaMap}
                  onChange={(e) => this.updateShortcut('toggleFormulaMap', e.target.value)}
                  className="shortcut-input"
                />
              </label>
            </div>

            <div className="shortcut-item">
              <label>
                <span className="shortcut-label">Trace Dependents</span>
                <input
                  type="text"
                  value={shortcuts.traceDependents}
                  onChange={(e) => this.updateShortcut('traceDependents', e.target.value)}
                  className="shortcut-input"
                />
              </label>
            </div>

            <div className="shortcut-item">
              <label>
                <span className="shortcut-label">Calculation Flow</span>
                <input
                  type="text"
                  value={shortcuts.calculationFlow}
                  onChange={(e) => this.updateShortcut('calculationFlow', e.target.value)}
                  className="shortcut-input"
                />
              </label>
            </div>
          </div>

          <button className="btn btn-secondary" onClick={this.resetShortcuts}>
            Reset to Defaults
          </button>
        </div>

        <div className="settings-section">
          <h3>Formula Explorer Options</h3>

          <label className="checkbox-label">
            <input
              type="checkbox"
              checked={explorerOptions.closeOnEscape}
              onChange={(e) =>
                this.setState({
                  explorerOptions: {
                    ...explorerOptions,
                    closeOnEscape: e.target.checked
                  }
                })
              }
            />
            Close window on Escape (and navigate back automatically)
          </label>

          <label className="checkbox-label">
            <input
              type="checkbox"
              checked={explorerOptions.autoNavigateBack}
              onChange={(e) =>
                this.setState({
                  explorerOptions: {
                    ...explorerOptions,
                    autoNavigateBack: e.target.checked
                  }
                })
              }
            />
            Keep selection on Enter (don't navigate back)
          </label>
        </div>

        <div className="settings-section">
          <h3>Navigation History</h3>
          <p className="section-description">
            Numeriq keeps track of up to 100 cell explorations for easy navigation
          </p>

          <div className="history-info">
            <p>
              Current history size: {KeyboardShortcutManager.getHistory().length} / 100
            </p>
            <p>
              Current position: {KeyboardShortcutManager.getCurrentHistoryIndex() + 1}
            </p>
          </div>

          <button className="btn btn-secondary" onClick={this.clearNavigationHistory}>
            Clear Navigation History
          </button>
        </div>

        <div className="settings-section">
          <h3>About Numeriq</h3>
          <div className="about-info">
            <p><strong>Version:</strong> 1.0.0</p>
            <p><strong>Description:</strong> Advanced Excel formula exploration and analysis tool</p>
            <p>
              Numeriq helps you understand complex spreadsheets by visualizing formula structures,
              comparing workbooks, mapping formula patterns, and analyzing calculation flows.
            </p>
          </div>
        </div>
      </div>
    );
  }
}
