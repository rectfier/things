import * as React from 'react';
import { createContext, useContext } from 'react';
import { SPFI } from "@pnp/sp";

export interface IPnPjsContext {
  sp: SPFI;
}

const PnPjsContext = createContext<IPnPjsContext>(null as any);

export const PnPjsProvider: React.FC<{ sp: SPFI; children?: React.ReactNode }> = ({ children, sp }) => {
  return (
    <PnPjsContext.Provider value={{ sp }}>
      {children}
    </PnPjsContext.Provider>
  );
};

export const usePnPjs = () => {
  const context = useContext(PnPjsContext);
  if (!context) {
    throw new Error("usePnPjs must be used within a PnPjsProvider");
  }
  return context;
};
