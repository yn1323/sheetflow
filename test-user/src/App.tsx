import { useState } from 'react'
import reactLogo from './assets/react.svg'
import viteLogo from '/vite.svg'
import './App.css'

import { createWorkbook } from 'xlkit';

function App() {
  const [count, setCount] = useState(0)

  const download = () => {
    excel()
  }

  return (
    <>
      <div>
        <a href="https://vite.dev" target="_blank">
          <img src={viteLogo} className="logo" alt="Vite logo" />
        </a>
        <a href="https://react.dev" target="_blank">
          <img src={reactLogo} className="logo react" alt="React logo" />
        </a>
      </div>
      <h1>Vite + React</h1>
      <div className="card">
        <button onClick={() => setCount((count) => count + 1)}>
          count is {count}
        </button>
        <p>
          Edit <code>src/App.tsx</code> and save to test HMR
        </p>
      </div>
      <p className="read-the-docs">
        Click on the Vite and React logos to learn more
      </p>
      <button onClick={download}>Download Excel</button>
    </>
  )
}

async function excel() {
  console.log("Generating Excel file...")
  
  // データ定義
  const fontData = [
    { label: 'Normal', text: 'The quick brown fox' },
    { label: 'Bold', text: 'The quick brown fox' },
    { label: 'Italic', text: 'The quick brown fox' },
    { label: 'Underline', text: 'The quick brown fox' },
    { label: 'Strike', text: 'The quick brown fox' },
    { label: 'Size 14', text: 'The quick brown fox' },
    { label: 'Size 18', text: 'The quick brown fox' },
    { label: 'Color Red', text: 'The quick brown fox' },
    { label: 'Color Blue', text: 'The quick brown fox' },
  ];
  
  await createWorkbook().addSheet({
    name: 'Fonts',
    headers: [
      { key: 'label', label: 'Style', width: 15 },
      { 
        key: 'text', 
        label: 'Preview', 
        width: 30,
        style: (val, row) => {
          if (row.label === 'Bold') return { font: { bold: true } };
          if (row.label === 'Italic') return { font: { italic: true } };
          if (row.label === 'Underline') return { font: { underline: true } };
          if (row.label === 'Strike') return { font: { strike: true } };
          if (row.label === 'Size 14') return { font: { size: 14 } };
          if (row.label === 'Size 18') return { font: { size: 18 } };
          if (row.label === 'Color Red') return { font: { color: '#FF0000' } };
          if (row.label === 'Color Blue') return { font: { color: '#0000FF' } };
          return {};
        }
      }
    ],
    rows: fontData
  }).download('font-styles.xlsx');
  
  console.log('Done! ✨');
}

export default App
