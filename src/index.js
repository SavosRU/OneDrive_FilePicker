import React from 'react';
import ReactDOM from 'react-dom';

import OneDriveFilePicker from './onedrive_filepicker.jsx';

function App() {
  return <OneDriveFilePicker />;
}

const rootElement = document.getElementById('root');
ReactDOM.render(<App />, rootElement);
