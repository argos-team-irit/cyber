import { useEffect, useState } from 'react';
import { HeaderHero } from './HeaderHero';
import { useLanguage } from '../LanguageContext';
import { translations } from '../translations';
import teamData from '../data.json';
import { read, utils } from 'xlsx';
import { GraduationCap, Loader2 } from 'lucide-react';

interface SupervisionItem {
  name: string;
  status: string;
  description: string;
  level: string;
}

export const SupervisionPage = () => {
  const { language } = useLanguage();
  const t = translations[language];
  const [items, setItems] = useState<SupervisionItem[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(false);

  useEffect(() => {
    const fetchSupervision = async () => {
      if (!teamData.supervisionXlsxUrl) {
        setLoading(false);
        return;
      }
      try {
        const proxyUrl = `/api/proxy-xlsx?url=${encodeURIComponent(teamData.supervisionXlsxUrl)}`;
        const response = await fetch(proxyUrl);
        if (!response.ok) {
          throw new Error('Failed to retrieve spreadsheet');
        }
        
        const buffer = await response.arrayBuffer();
        const workbook = read(buffer, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        const rows = utils.sheet_to_json<any>(worksheet);
        
        const parsedItems: SupervisionItem[] = rows.map((row) => {
          // Fallback parsing, trying to find matching keys
          const getVal = (keys: string[]) => {
            const rowKeys = Object.keys(row);
            for (const k of keys) {
              const matchedKey = rowKeys.find(rk => rk.toLowerCase().trim() === k.toLowerCase());
              if (matchedKey && row[matchedKey]) return String(row[matchedKey]).trim();
            }
            return '';
          };
          
          return {
            name: getVal(['name', 'nom']),
            status: getVal(['status', 'statut']),
            description: getVal(['description', 'desc']),
            level: getVal(['level', 'niveau', 'type']),
          };
        }).filter(item => item.name || item.description); // filter out empty rows
        
        setItems(parsedItems);
        setError(false);
      } catch (err) {
        console.error("Supervision fetch error:", err);
        setError(true);
      } finally {
        setLoading(false);
      }
    };
    fetchSupervision();
  }, []);

  return (
    <>
      <HeaderHero isHome={false} />
      
      <div className="bg-slate-50 min-h-[60vh] py-24 border-t border-slate-200">
        <div className="max-w-5xl mx-auto px-6">
          <div className="mb-14 text-center">
            <h2 className="flex items-center justify-center text-4xl font-bold tracking-tight text-slate-900 mb-4 font-serif">
              <GraduationCap className="w-10 h-10 mr-4 text-blue-600" />
              {language === 'en' ? 'Supervising' : 'Supervision'}
            </h2>
            <p className="text-lg text-slate-600 max-w-2xl mx-auto">
              {language === 'en' 
                ? 'Overview of supervised students and research projects.' 
                : 'Aperçu des étudiants supervisés et des projets de recherche.'}
            </p>
          </div>

          {loading && (
            <div className="flex justify-center items-center py-12 text-slate-500">
              <Loader2 className="w-8 h-8 animate-spin mr-3" />
              <span className="text-lg">{t.publications.loading}</span>
            </div>
          )}

          {error && (
            <div className="bg-red-50 text-red-600 p-6 rounded-xl border border-red-100 text-center max-w-lg mx-auto">
              {language === 'en' 
                ? 'Could not load the supervision list at this time.' 
                : 'Impossible de charger la liste de supervision pour le moment.'}
            </div>
          )}

          {!loading && !error && items.length === 0 && (
            <div className="text-center py-12 bg-white rounded-xl border border-slate-200 text-slate-500">
              {language === 'en' ? 'No records found.' : 'Aucun enregistrement trouvé.'}
            </div>
          )}

          {!loading && !error && items.length > 0 && (
            <div className="space-y-20">
              {(() => {
                const groupedByLevel: Record<string, SupervisionItem[]> = {};
                items.forEach(item => {
                  const levelRaw = item.level || 'Other';
                  let levelGroup = 'Other';
                  const l = levelRaw.toLowerCase();
                  if (l.includes('phd') || l.includes('doctora') || l.includes('thèse') || l.includes('these')) {
                    levelGroup = language === 'en' ? 'PhD Students' : 'Doctorants';
                  } else if (l.includes('master') || l.includes('m1') || l.includes('m2')) {
                    levelGroup = language === 'en' ? 'Master Students' : 'Étudiants en Master';
                  } else if (l.includes('bachelor') || l.includes('licence') || l.includes('l3')) {
                    levelGroup = language === 'en' ? 'Bachelor Students' : 'Étudiants en Licence';
                  } else {
                    levelGroup = levelRaw;
                  }
                  if (!groupedByLevel[levelGroup]) groupedByLevel[levelGroup] = [];
                  groupedByLevel[levelGroup].push(item);
                });

                // Preferred order of keys
                const orderedLevels = [
                  language === 'en' ? 'PhD Students' : 'Doctorants',
                  language === 'en' ? 'Master Students' : 'Étudiants en Master',
                  language === 'en' ? 'Bachelor Students' : 'Étudiants en Licence'
                ];
                
                const existingLevels = Object.keys(groupedByLevel).sort((a, b) => {
                  const idxA = orderedLevels.indexOf(a);
                  const idxB = orderedLevels.indexOf(b);
                  if (idxA !== -1 && idxB !== -1) return idxA - idxB;
                  if (idxA !== -1) return -1;
                  if (idxB !== -1) return 1;
                  return a.localeCompare(b);
                });

                return existingLevels.map((levelName) => {
                  const levelItems = groupedByLevel[levelName];
                  
                  const ongoingItems = levelItems.filter(item => {
                    const s = item.status.toLowerCase();
                    return !s.includes('completed') && !s.includes('terminé') && !s.includes('soutenue');
                  });
                  
                  const completedItems = levelItems.filter(item => {
                    const s = item.status.toLowerCase();
                    return s.includes('completed') || s.includes('terminé') || s.includes('soutenue');
                  });

                  return (
                    <div key={levelName}>
                      <h3 className="text-3xl font-bold text-slate-900 mb-8 font-serif border-b-2 pb-3 border-blue-100">
                        {levelName}
                      </h3>
                      
                      <div className="space-y-12 pl-2">
                        {/* Ongoing Section */}
                        {ongoingItems.length > 0 && (
                          <div>
                            <h4 className="text-xl font-bold text-slate-800 mb-5 font-serif border-b pb-2 border-slate-100">
                              {language === 'en' ? 'Ongoing' : 'En cours'}
                            </h4>
                            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                              {ongoingItems.map((item, idx) => (
                                <div key={idx} className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm hover:shadow-md transition-shadow">
                                  <h5 className="text-lg font-bold text-slate-900 font-serif mb-1">{item.name || 'Unknown'}</h5>
                                  <span className="inline-block px-3 py-1 bg-green-50 text-green-700 text-xs font-semibold rounded-full mb-3 uppercase tracking-wider">
                                    {item.status || (language === 'en' ? 'Ongoing' : 'En cours')}
                                  </span>
                                  <p className="text-slate-600 leading-relaxed text-sm text-justify">
                                    {item.description}
                                  </p>
                                </div>
                              ))}
                            </div>
                          </div>
                        )}

                        {/* Completed Section */}
                        {completedItems.length > 0 && (
                          <div>
                            <h4 className="text-xl font-bold text-slate-800 mb-5 font-serif border-b pb-2 border-slate-100">
                              {language === 'en' ? 'Completed' : 'Terminé'}
                            </h4>
                            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                              {completedItems.map((item, idx) => (
                                <div key={idx} className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm hover:shadow-md transition-shadow opacity-80">
                                  <h5 className="text-lg font-bold text-slate-900 font-serif mb-1">{item.name || 'Unknown'}</h5>
                                  <span className="inline-block px-3 py-1 bg-slate-100 text-slate-700 text-xs font-semibold rounded-full mb-3 uppercase tracking-wider">
                                    {item.status}
                                  </span>
                                  <p className="text-slate-600 leading-relaxed text-sm text-justify">
                                    {item.description}
                                  </p>
                                </div>
                              ))}
                            </div>
                          </div>
                        )}
                      </div>
                    </div>
                  );
                });
              })()}
            </div>
          )}
        </div>
      </div>
    </>
  );
};
