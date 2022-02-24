import { Agenda, Login, FileList } from '@microsoft/mgt-react';
import React, { useState, useEffect } from 'react';
//import './App.css';

import { Providers, ProviderState } from '@microsoft/mgt-element';

function useIsSignedIn(): [boolean] {
  const [isSignedIn, setIsSignedIn] = useState(false);

  useEffect(() => {
    const updateState = () => {
      const provider = Providers.globalProvider;
      setIsSignedIn(provider && provider.state === ProviderState.SignedIn);
    };

    Providers.onProviderUpdated(updateState);
    updateState();

    return () => {
      Providers.removeProviderUpdatedListener(updateState);
    }
  }, []);

  return [isSignedIn];
}

function App() {
  const [isSignedIn] = useIsSignedIn();

  return (
    <div className="App">
      <header>
        <Login />
      </header>
      {/*
      <div>
        {isSignedIn &&
          <Agenda />
          }
      </div> */}
      <div>
        {isSignedIn &&
          <FileList driveId="b!mKw3q1anF0C5DyDiqHKMr8iJr_oIRjlGl4854HhHtho07AdbOeaLT5rMH83yt89B" 
          itemPath="/" enableFileUpload></FileList>
          }
      </div> 
    </div>
  );
}

export default App;
