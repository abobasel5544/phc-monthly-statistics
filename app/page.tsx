'use client';
import { useState } from 'react';
import * as XLSX from 'xlsx';
import { modules } from '../lib/templates';

export default function Page() {
  const [status, setStatus] = useState<string>('');

  const processFile = (e: React.ChangeEvent<HTMLInputElement>, moduleId: string) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const wb = XLSX.read(evt.target?.result, { type: 'binary' });
      const mod = modules.find(m => m.id === moduleId);
      const ws = wb.Sheets[mod?.sheetName || wb.SheetNames[0]];

      // معالجة البيانات بدءاً من الصف المحدد في القالب
      const jsonData = XLSX.utils.sheet_to_json(ws, { 
        header: 1, 
        range: (mod?.firstMonthRow || 3) - 1 
      });

      console.log(`بيانات ${mod?.title}:`, jsonData);
      setStatus(`تم بنجاح رفع ومعالجة بيانات: ${mod?.title}`);
    };
    reader.readAsBinaryString(file);
  };

  return (
    <div className="p-8 space-y-6">
      <h1 className="text-2xl font-bold">لوحة رفع البيانات الصحية</h1>
      {modules.map(mod => (
        <div key={mod.id} className="border p-4 rounded">
          <label className="block mb-2 font-semibold">{mod.title}</label>
          <input type="file" onChange={(e) => processFile(e, mod.id)} />
        </div>
      ))}
      <div className="p-4 bg-green-100 text-green-800 rounded">{status}</div>
    </div>
  );
}
