import React, { useEffect, useState } from 'react';
import { createRoot } from 'react-dom/client';
import { Moon, Sun } from 'lucide-react';
import './styles.css';

function App() {
  const [dark, setDark] = useState(() => localStorage.getItem('at-dark') === '1');
  const [boot, setBoot] = useState(true);

  useEffect(() => {
    const timer = setTimeout(() => setBoot(false), 900);
    return () => clearTimeout(timer);
  }, []);

  const toggleTheme = () => {
    const next = !dark;
    setDark(next);
    localStorage.setItem('at-dark', next ? '1' : '0');
  };

  return (
    <div className={dark ? 'app dark' : 'app'}>
      {boot && (
        <div className="siteIntro">
          <div className="introBrand">Abdel-hamid Zahran</div>
          <div className="introLine" />
          <h1>أدوات المحاسب</h1>
        </div>
      )}

      <header className="header">
        <div className="brand">
          <div className="azLogo" aria-label="AZ">AZ</div>
          <div className="brandText">
            <strong>Abdel-hamid Zahran</strong>
            <span>أدوات المحاسب</span>
          </div>
        </div>

        <button className="themeButton" onClick={toggleTheme} aria-label="تغيير المظهر">
          {dark ? <Sun size={20} /> : <Moon size={20} />}
        </button>
      </header>

      <main className="emptyHome">
        <div className="emptyIcon">AZ</div>
        <h1>أدوات المحاسب</h1>
        <p>الموقع جاهز لبناء الأدوات من البداية.</p>
        <span>لا توجد أدوات PDF أو أدوات قديمة في النسخة الجديدة.</span>
      </main>

      <footer>
        <b>أدوات المحاسب</b>
        <span>Accountant Abdel-hamid Zahran</span>
      </footer>
    </div>
  );
}

createRoot(document.getElementById('root')).render(<App />);
