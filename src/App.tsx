import { LanguageProvider } from './LanguageContext';
import { HeaderHero } from './components/HeaderHero';
import { AboutProjects } from './components/AboutProjects';
import { Team } from './components/Team';
import { PublicationsSections } from './components/PublicationsSections';
import { Footer } from './components/Footer';
import { SupervisionPage } from './components/SupervisionPage';
import { BrowserRouter, Routes, Route } from 'react-router-dom';

function Home() {
  return (
    <>
      <HeaderHero isHome={true} />
      <AboutProjects />
      <Team />
      <PublicationsSections />
    </>
  );
}

export default function App() {
  return (
    <LanguageProvider>
      <BrowserRouter>
        <div className="min-h-screen bg-white font-sans text-slate-900 selection:bg-blue-100 selection:text-blue-900">
          <main>
            <Routes>
              <Route path="/" element={<Home />} />
              <Route path="/supervision" element={<SupervisionPage />} />
            </Routes>
          </main>
          <Footer />
        </div>
      </BrowserRouter>
    </LanguageProvider>
  );
}
