import { useEffect } from 'react';
import * as obj from '../common/module';
import logo from '../logo.svg';
import '../common/other';
import './App.css';

setTimeout(() => {
  console.warn(obj.age)
}, 1000)

function App() {
  useEffect(() => {
    console.log(obj.age)
  }, [])
  return (
    <div className="App">
      <header className="App-header">
        <img src={logo} className="App-logo" alt="logo" />
        <p>
          Edit <code>src/App.js</code> and save to reload.
        </p>
        <a
          className="App-link"
          href="https://reactjs.org"
          target="_blank"
          rel="noopener noreferrer"
        >
          Learn React
        </a>
      </header>
    </div>
  );
}

export default App;
