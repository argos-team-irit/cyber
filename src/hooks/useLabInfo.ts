import { useEffect, useState } from 'react';
import { read, utils } from 'xlsx';
import data from '../data.json';
import { translations } from '../translations';

export const useLabInfo = (language: 'en' | 'fr') => {
  const t = translations[language];
  const [info, setInfo] = useState({
    short: 'CS',
    labEn: translations.en.hero.title,
    labFr: translations.fr.hero.title,
    descEn: translations.en.hero.subtitle,
    descFr: translations.fr.hero.subtitle,
    affiliation: translations[language].footer.affiliation,
    address: translations[language].footer.address,
    contact: translations[language].footer.contact,
    logo: '',
    visionEn: '',
    visionFr: '',
    contentEn: '',
    contentFr: '',
  });

  useEffect(() => {
    if (!data.infoXlsxUrl) return;

    fetch(`/api/proxy-xlsx?url=${encodeURIComponent(data.infoXlsxUrl)}`)
      .then(res => res.arrayBuffer())
      .then(buffer => {
        const workbook = read(buffer, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        
        // Use header: 1 to get an array of arrays, representing rows.
        const rows = utils.sheet_to_json(worksheet, { header: 1 });
        
        const obj: Record<string, string> = {};
        rows.forEach((row: any) => {
          if (Array.isArray(row) && row.length >= 2) {
            const key = String(row[0]).toLowerCase().trim();
            const value = String(row[1]).trim();
            if (key) obj[key] = value;
          }
        });

        const findValue = (keys: string[]) => {
          for (const key of keys) {
             if (obj[key] && obj[key] !== '') return obj[key];
          }
          return null;
        };

        if (Object.keys(obj).length > 0) {
          setInfo(prev => ({
            ...prev,
            short: findValue(['short', 'abbreviation']) || prev.short,
            labEn: findValue(['laben', 'lab en', 'lab_en', 'lab-en', 'title en', 'lab', 'team', 'title']) || prev.labEn,
            labFr: findValue(['labfr', 'lab fr', 'lab_fr', 'lab-fr', 'title fr', 'lab', 'team', 'title']) || prev.labFr,
            descEn: findValue(['descen', 'description en', 'desc en', 'description_en', 'description', 'desc', 'subtitle']) || prev.descEn,
            descFr: findValue(['descfr', 'description fr', 'desc fr', 'description_fr', 'description', 'desc', 'subtitle']) || prev.descFr,
            logo: (() => {
              const val = findValue(['logo', 'image', 'photo', 'icon']);
              if (val) return val.startsWith('http') || val.startsWith('data:') ? val : `https://${val}`;
              return prev.logo;
            })(),
            affiliation: findValue(['affiliation']) || prev.affiliation,
            address: findValue(['address', 'adresse']) || prev.address,
            contact: findValue(['contact', 'email', 'courriel']) || prev.contact,
            visionEn: findValue(['visionen', 'vision en', 'vision_en', 'vision-en', 'vision en']) || prev.visionEn,
            visionFr: findValue(['visionfr', 'vision fr', 'vision_fr', 'vision-fr', 'vision fr']) || prev.visionFr,
            contentEn: findValue(['contenten', 'content en', 'content_en', 'content-en', 'content en']) || prev.contentEn,
            contentFr: findValue(['contentfr', 'content fr', 'content_fr', 'content-fr', 'content fr']) || prev.contentFr,
          }));
        }
      })
      .catch(e => console.error("Failed to load info excel:", e));
  }, []);

  return {
    ...info,
    currentLab: language === 'fr' ? info.labFr : info.labEn,
    currentDesc: language === 'fr' ? info.descFr : info.descEn,
  };
};
