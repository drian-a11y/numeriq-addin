import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { App } from './components/App';
import './taskpane.css';

/* global document, Office */

Office.onReady(() => {
  ReactDOM.render(<App />, document.getElementById('container'));
});
