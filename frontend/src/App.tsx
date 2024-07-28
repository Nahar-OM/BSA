import { Button } from '@nextui-org/button';
import React, { useState, useEffect } from 'react';
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";

const BSAProcessor: React.FC = () => {
  const [isProcessing, setIsProcessing] = useState(false);
  const [progress, setProgress] = useState(0);
  const [downloadUrl, setDownloadUrl] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [selectedFolder, setSelectedFolder] = useState<string | null>(null);

  const folders = [
    { value: 'All_Bank_BS', label: 'All Bank BS' },
    { value: 'Amal_Transports_and_Travels', label: 'Amal Transports and Travels' },
    { value: 'ASP_Traders', label: 'ASP Traders' },
    { value: 'Ayynar_Travels', label: 'Ayynar Travels' },
    { value: 'Bank_Statement_FY_2020-2021', label: 'Bank Statement FY 2020-2021' },
    { value: 'CAROL', label: 'CAROL' },
    { value: 'CUB', label: 'CUB' },
    { value: 'DIA_SKILLS_DEVELOPMENT', label: 'DIA SKILLS DEVELOPMENT' },
    { value: 'East_Sun_Logistics', label: 'East Sun Logistics' },
    { value: 'EXPERT_CAPITAL_SERVICES', label: 'EXPERT CAPITAL SERVICES' },
    { value: 'GSG_Logistics', label: 'GSG Logistics' },
    { value: 'Jaldeepan', label: 'Jaldeepan' },
    { value: 'Kriya', label: 'Kriya' },
    { value: 'LANDCRAFT-DEVELOPERS', label: 'LANDCRAFT DEVELOPERS' },
    { value: 'LANDCRAFT-RECREATIONS', label: 'LANDCRAFT RECREATIONS' },
    { value: 'Lavanya', label: 'Lavanya' },
    { value: 'LS_Trading', label: 'LS Trading' },
    { value: 'MAA_TARINI_Agencies', label: 'MAA TARINI Agencies' },
    { value: 'MERIDIAN_CITY_PROJECT_PRIVATE', label: 'MERIDIAN CITY PROJECT PRIVATE' },
    { value: 'MS_Kriya', label: 'MS Kriya' },
    { value: 'MS_Reeshma', label: 'MS Reeshma' },
    { value: 'MVH_SOLUTIONS', label: 'MVH SOLUTIONS' },
    { value: 'N_Karthick', label: 'N Karthick' },
    { value: 'Others', label: 'Others' },
    { value: 'PDF_Samples', label: 'PDF Samples' },
    { value: 'Prag_Yuga', label: 'Prag Yuga' },
    { value: 'S_Bharathiraja', label: 'S Bharathiraja' },
    { value: 'STG_Tours_and_Travels', label: 'STG Tours and Travels' },
    { value: 'TAMILNADU_SANITARY_STORES', label: 'TAMILNADU SANITARY STORES' },
  ];

  useEffect(() => {
    setDownloadUrl(null);
  }, [selectedFolder]);

  const startBSAProcess = async () => {
    if (!selectedFolder) {
      setError('Please select a folder.');
      return;
    }

    setIsProcessing(true);
    setProgress(0);
    setDownloadUrl(null);
    setError(null);

    const eventSource = new EventSource(`http://localhost:3000/run-bsa?folder=${selectedFolder}`);

    eventSource.onmessage = (event) => {
      const data = event.data;

      if (data.startsWith('BSA process started')) {
        setProgress(10);
      } else if (data.startsWith('BSA process completed')) {
        setProgress(90);
      } else if (data.startsWith('Download URL:')) {
        const url = data.split(': ')[1];
        setDownloadUrl(url);
        setProgress(100);
        setIsProcessing(false);
        eventSource.close();
      } else if (data.startsWith('Error:')) {
        setError(data);
        setIsProcessing(false);
        eventSource.close();
      }
    };

    eventSource.onerror = (error) => {
      setError(`An error occurred during the BSA process: ${JSON.stringify(error)}`);
      setIsProcessing(false);
      eventSource.close();
    };
  };

  return (
    <div className='mx-auto hero-section phone:px-5'>
      <div className='background-lines'>
        <div className='line'></div>
        <div className='line'></div>
        <div className='line'></div>
      </div>
      <div className=''>
        <div className='flex mx-auto badge max-w-max'>
          <img src='/icons/star.svg' alt='star' className='w-6 h-6 my-auto'/>
          <p className='text-sm font-medium gradient-text text-nowrap'>Advanced analysis</p>
        </div>
        <div className='pt-5 font-semibold leading-normal tracking-tighter text-center phone:text-2xl tablet:text-6xl'>
          <h1 className='tracking-tighter lg:pb-2 gradient-text'>AI-Powered</h1>
          <span className='italic font-normal tracking-tight font-instrument-serif gradient-text'>Bank Statement Analyser</span>
        </div>
        <p className='tablet:w-[80%] mx-auto text-center tracking-tight phone:text-sm tablet:text-base mt-7 text-wrap phone:px-5 text-gray-400'>
          Effortlessly understand your spending, track income, and achieve financial clarity with our Bank Statement Analyzer
        </p>
        <div className="max-w-md mx-auto text-center mt-7">
          <Select onValueChange={(value: React.SetStateAction<string | null>) => setSelectedFolder(value)}>
            <SelectTrigger className='z-50 px-4 py-2 mt-5 text-sm text-gray-300 border-none rounded-lg outline-none card-cover disabled:bg-gray-400 disabled:cursor-not-allowed'>
              <SelectValue placeholder="Select Folder" />
            </SelectTrigger>
            <SelectContent side='bottom' className='border-none card-cover'>
              {folders.map((folder) => (
                <SelectItem key={folder.value} value={folder.value} className='z-10 cursor-pointer gradient-text hover:bg-white/50'>
                  {folder.label}
                </SelectItem>
              ))}
            </SelectContent>
          </Select>

          {!isProcessing && !downloadUrl && (
            <Button
              onClick={startBSAProcess}
              disabled={isProcessing}
              className="z-50 px-4 py-2 mt-5 text-sm text-gray-300 rounded-lg card-cover disabled:bg-gray-400 disabled:cursor-not-allowed"
            >
              Start BSA Process
            </Button>
          )}

          {isProcessing && (
            <div className="w-full h-5 mt-5 rounded-lg progress-bar card-cover">
              <div className="h-full transition-all duration-500 rounded-lg progress card-cover" style={{ width: `${progress}%` }}></div>
            </div>
          )}

          {downloadUrl && (
            <Button className="z-50 px-5 py-3 mt-5 text-sm text-gray-300 rounded-lg card-cover disabled:bg-gray-400 disabled:cursor-not-allowed">
              <a href={downloadUrl} download>
                Download Result
              </a>
            </Button>
          )}

          {error && <p className="mt-5 text-sm gradient-text error">{error}</p>}
        </div>
        <div className="flex items-center justify-center p-4 mt-12">
          <div className="flex mr-4 -space-x-3">
            <div className="w-8 h-8 bg-purple-500 border-2 border-gray-900 rounded-full"></div>
            <div className="w-8 h-8 bg-green-500 border-2 border-gray-900 rounded-full"></div>
            <div className="w-8 h-8 bg-orange-500 border-2 border-gray-900 rounded-full"></div>
            <div className="flex items-center justify-center w-8 h-8 bg-white border-2 border-gray-900 rounded-full">
              <div className="w-4 h-4 bg-gray-900 rounded-full"></div>
            </div>
          </div>
          <span className="text-sm font-medium text-gray-600">100,000+ users</span>
        </div>
      </div>
    </div>
  );
};

export default BSAProcessor;
