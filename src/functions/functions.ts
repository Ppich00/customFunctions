/* global clearInterval, console, CustomFunctions, setInterval */
import * as platform from "@office-platform/office-web-service-client";
/**
 * Adds two numbers.
 * @customfunction
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */
export function add(first: number, second: number): number {
  return first + second;
}

/**
 * Displays the current time once a second.
 * @customfunction
 * @param invocation Custom function handler
 */
export function clock(invocation: CustomFunctions.StreamingInvocation<string>): void {
  const timer = setInterval(() => {
    const time = currentTime();
    invocation.setResult(time);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Returns the current time.
 * @returns String with the current time formatted for the current locale.
 */
export function currentTime(): string {
  return new Date().toLocaleTimeString();
}

/**
 * Increments a value once a second.
 * @customfunction
 * @param incrementBy Amount to increment
 * @param invocation Custom function handler
 */
export function increment(incrementBy: number, invocation: CustomFunctions.StreamingInvocation<number>): void {
  let result = 0;
  const timer = setInterval(() => {
    result += incrementBy;
    invocation.setResult(result);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Writes a message to console.log().
 * @customfunction LOG
 * @param message String to write.
 * @returns String to write.
 */
export function logMessage(message: string): string {
  console.log(message);

  return message;
}

// =RDP.Data("CNY=;GBP=;CHF=;NZD=;SEK=","CF_BID;CF_ASK;CF_NETCHNG","CH=FD RH=IN")
let fullTable: string[][] = [[]];
/**
 * data
 * @customfunction data
 * @param invocation Custom function handler
 */
export async function data(invocation?: CustomFunctions.StreamingInvocation<string[][]>) {
  platform.configure({
    platform: "beta",
  });
  const session = platform.createSession({ autoOpen: false });
  debugger;
  const token =
    "eyJ0eXAiOiJhdCtqd3QiLCJhbGciOiJSUzI1NiIsImtpZCI6InRhbTg2amltMDY4dTh0LWpPU1c3ZWhyTUx6VW1INVVhOS1JOGRtTzA3Zm8ifQ.eyJkYXRhIjoie1wiY2lwaGVydGV4dFwiOlwiTGdqajgtUE5DdXY3SVhnazNLaTFTdEhNQmxPa2VDemRBaDNMa2tkNmItQ0lwb0x3NFkwQTdpMGphdGppTzludnVMS3dhY3JPbWpfNGdTbHFhMjlLTFQ2NDVhU0RoTi1xRU1JMDRwR09PQzRzYTNTamJPSmtSZFU4YXBiSEhSd29nSHFhYTAydkl4cDNaWGw4NmZzWlpUeHMtVG9yNlMycHVicUZmcVhqdHVTWmxITnhHOHRxNWJWT0FtSTlZaVRmMlN6SUpQZmg2M1ExanBSZTVGLVVzdmtsTFREV2w0R1JHSnVTa1J6V09tTm1GLUVyUkRqVUVlemFTT3Y5Skd6U1NFdUgzR3R6LXE0YTRVUEhUR2NCdGdGNXp1MEJsZGowa1BIbU9nakl0MDg3V2ZyN2VWZHhJbm9GaVB4QVNZNGcxbWxhUjBNOTlxNEtxU2FhaVpPQjMtNWRYd2Vad0tLY2Q4Qzl6TDJ0QVBReXYxZ003enhjVmR2ZVh5VlNCYjlSdUFCdDVMbkR6ajBIaS1VeW5tY01HcXdaMXBHeWlsY1pyT1dPMVBHQ0Vsclp5MFlyeHdlYW5BQmc5dEVmeS0yUGtMNUN4cU42UXZ5ck5hbk1HOXlhaHl0czVsVFRNR0VsazN2QVlHWV84MDMxNXlfQTZtd0ZiSkcyRWNXT0hpdDVCbjVVc2piNXNXX0JVNkFybVVSQm1HMmdwOUhXVVduV1Rpc2dCZzc4Y250dGl1aDlrcW0xX2xvZGNDazZWcjZScXdvbkVoZExJOV9saWZKZG42M3lEZ1pOSEwwMWFFTmQ1eFZTby1wbzgzQndYbVdhNzctZ1hRR0hoeXF6ZEowWjkwOExTV3ZTRUp4Z1lEdDdmNUR6ek05M3d5dGVzYkpYMC1RQUpxRUNLYnU2dVp1UlRhYW5FYk9raFJOTnBOQ2gxVXoydzZaVWNVY3ExNW8yOXVCTGhUVFNfRWNkTzl0REZGQUtpQnBhNFAxS1A5SWR3SXprNjBJZ0xKNVdKb1hYMzQ2VXB6UkZZWWE5RkF1bGF2NnloVm5CVmE0RXNsWFREZ01nMmNZVldEVzZOT1JIbnhrMjN6aDZYVldwRVdpNDhKcHA5Q2tcIixcIml2XCI6XCI5enJlQ0MzV0NxRnVFSnprXCIsXCJwcm90ZWN0ZWRcIjpcImV5SmhiR2NpT2lKQlYxTmZSVTVEWDFORVMxOUJNalUySWl3aVpXNWpJam9pUVRJMU5rZERUU0lzSW5wcGNDSTZJa1JGUmlKOVwiLFwicmVjaXBpZW50c1wiOlt7XCJlbmNyeXB0ZWRfa2V5XCI6XCJBUUlCQUhqOW9sZDI5RXB3VU1hZTZCSjRRVWVrOGZMQTlxQ2t3YlBKUFF1QV9tMEdwUUdTSGFPUm01Q0w0TktaTXN3eGtkaWZBQUFBZmpCOEJna3Foa2lHOXcwQkJ3YWdiekJ0QWdFQU1HZ0dDU3FHU0liM0RRRUhBVEFlQmdsZ2hrZ0JaUU1FQVM0d0VRUU1uVEdqUWM4NEo4VmRuWW1KQWdFUWdEdlZnTTkycDhtNE4ycVVxd0I2MG9RMnBiM3VINE1SMFJPYkJKeWR4SmtwbW5kNWZTbkp5NElDajJub2E3d2kzcUVpYWkwN2hqX0h4cmMzNlFcIixcImhlYWRlclwiOntcImtpZFwiOlwiYXJuOmF3czprbXM6dXMtZWFzdC0xOjY1MzU1MTk3MDIxMDprZXkvNjUwOGQzM2YtMjU4My00YjNmLTlkZjItYmQwZTQ0YzUxZTQzXCJ9fSx7XCJlbmNyeXB0ZWRfa2V5XCI6XCJBUUlDQUhqbFYzZ2Vsa3EwbGJtSERySmxTcFBMaUtlNGozY1p4MkttdHd2MFR2cVlMZ0U3bHlrcXVTb1ZpWTNMZzdJTjlKeHZBQUFBZmpCOEJna3Foa2lHOXcwQkJ3YWdiekJ0QWdFQU1HZ0dDU3FHU0liM0RRRUhBVEFlQmdsZ2hrZ0JaUU1FQVM0d0VRUU1KNlNVSkhoMzJqTTlTRGctQWdFUWdEdG1iQXlMbm9kRlg1M1VHR3NWcGJmZ0xfVUM2dmpGcmZNQURrSHNzajRCUEpRb3lFX1R4WkdydzRIWFdzQUxZSHF6U3VzajgyUVpkd0FGSlFcIixcImhlYWRlclwiOntcImtpZFwiOlwiYXJuOmF3czprbXM6YXAtc291dGhlYXN0LTE6NjUzNTUxOTcwMjEwOmtleS82MmEzNjgzNS1iY2I1LTQ3ZjktYmIwNS03OTkyMDU3YzE1YzdcIn19LHtcImVuY3J5cHRlZF9rZXlcIjpcIkFRSUJBSGo5b2xkMjlFcHdVTWFlNkJKNFFVZWs4ZkxBOXFDa3diUEpQUXVBX20wR3BRR1NIYU9SbTVDTDROS1pNc3d4a2RpZkFBQUFmakI4QmdrcWhraUc5dzBCQndhZ2J6QnRBZ0VBTUdnR0NTcUdTSWIzRFFFSEFUQWVCZ2xnaGtnQlpRTUVBUzR3RVFRTW5UR2pRYzg0SjhWZG5ZbUpBZ0VRZ0R2VmdNOTJwOG00TjJxVXF3QjYwb1EycGIzdUg0TVIwUk9iQkp5ZHhKa3BtbmQ1ZlNuSnk0SUNqMm5vYTd3aTNxRWlhaTA3aGpfSHhyYzM2UVwiLFwiaGVhZGVyXCI6e1wia2lkXCI6XCJhcm46YXdzOmttczp1cy1lYXN0LTE6NjUzNTUxOTcwMjEwOmtleS82NTA4ZDMzZi0yNTgzLTRiM2YtOWRmMi1iZDBlNDRjNTFlNDNcIn19XSxcInRhZ1wiOlwiR1VoU3dTbVBDRkdoNjBjWFVwXzNKd1wifSIsInJzMSI6ImI1NjFjN2RhNDVlN2M1MDk1NzQ3ZTlkOTY4NjJlMTU5YTIwZmIyZDQiLCJhdWQiOiI4NDMxN2U1MjU3NGQ0YTliYjY1NTZjYzkyOTY4MmZmMWQwNDQ2NDlhIiwiaXNzIjoiaHR0cHM6Ly9pZGVudGl0eS5jaWFtLnJlZmluaXRpdi5jb20vYXBpL2lkZW50aXR5L3N0c19wcmVwcm9kIiwiZXhwIjoxNjUyNTk0MTc3LCJpYXQiOjE2NTI1OTM1Nzd9.t3kttbxpr4L_QxiCLn0Mm2dUXBsyx-FrHH_DtuAspHZGYBc2edCoJe2MRUy0NNirtruPqow5itqnqDElOaQCwPFAzKH3E1_vUyK_oNAyP3eSHkqgz7WoEdlx4PyxmPI9lIhoYxtUuJwYHgYwMRMhSnqEeZYk-N2r_tVU3G6zeIfmvNZqHPExVbr0AlgGS7Hic8Df-zB5_8wpGxaA6CNvQtmx0NbQTf5mXsFVR2W1e3YEA_IFTndMKbrehU8oB4JJi0-gSZFYm9RNI2ELsSi0Tnzuru5Fd8s2sZaLW9JqYX4guUH4f5whXNwUz3F34LG5pab42xSiJEJ2bNbWC6DTyQ";
  const opts = {
    token: `Bearer ${token}`,
    timeout: -1,
  };
  session
    .open(opts)
    .then((val) => {
      console.log("data value: ", val);
    })
    .catch((e) => {
      console.error("data error", e);
    })
    .finally(() => {
      console.log("data finally");
    });
  session.on("open", () => {
    const subscription = session.createSubscription({
      name: "RDP.Data",
      args: ["CNY=;GBP=;CHF=;NZD=;SEK=", "CF_BID;CF_ASK;CF_NETCHNG", "CH=FD RH=IN", null],
    });
    subscription.on("update", (functionResult: platform.FunctionResult) => {
      const functionResponse: IFunctionResponse = functionResult as IFunctionResponse;
      console.log(`function result: ${JSON.stringify(functionResult)}`);
      if (functionResponse.type === "snapshot") {
        const { showableTable, auditIdTable } = transformSnapshotResponse(functionResponse);
        const hasClickThrough: boolean = functionResponse.results.some((result) => result.clickThrough);
        if (fullTable.toString() !== showableTable.toString()) {
          fullTable = [...showableTable];

          invocation.setResult(fullTable);
        }
      } else {
        fullTable = [
          ...transformUpdateResponse(fullTable, functionResponse.results, functionResponse.callerValue as string),
        ];
        invocation.setResult(fullTable);
      }
    });
    setInterval(() => subscription.refresh(), 1000);

    setInterval(() => session.updateToken(opts.token), 240000);

    setTimeout(() => {
      session.pause();
    }, 240000);
  });
}

export const transformUpdateResponse = (
  fullTableData: string[][],
  updatedResults: IFunctionResultObject[],
  callerValue?: string
): string[][] => {
  let newTable: string[][] = [...fullTableData];
  updatedResults?.forEach((updatedResult: IFunctionResultObject): void => {
    newTable = [...mapUpdatedDataToTable(newTable, updatedResult)];
  });
  if (callerValue) {
    newTable[0][0] = callerValue;
  }
  return newTable;
};

export interface IFunctionResultObject {
  position: {
    row: number;
    column: number;
  };
  data: Array<(string | number | null | Record<string, unknown>)[]>;
  clickThrough?: Array<(string | null)[]>;
}
export interface IFunctionResponse {
  callerClickThrough?: string;
  callerValue: string | number | null;
  results?: IFunctionResultObject[];
  type: "update" | "snapshot";
}

const mapUpdatedDataToTable = (fullTableData: string[][], updatedResult: IFunctionResultObject): string[][] => {
  const replacedDataRowIndex: number = updatedResult.position.row;
  const replacedDataColIndex: number = updatedResult.position.column;
  let updatedRowCount = 0;
  let updatedColCount = 0;
  return fullTableData.map((rows: string[], rowIndex: number) => {
    let updateRows: string[] = [...rows];
    if (rowIndex >= replacedDataRowIndex && updatedRowCount < updatedResult.data.length) {
      const updatedRow: string[] = updatedResult.data[updatedRowCount] as string[];
      updateRows = rows.map((col: string, colIndex: number): string => {
        let updatedColValue: string = col;
        if (colIndex >= replacedDataColIndex && updatedColCount < updatedRow.length) {
          updatedColValue = updatedRow[updatedColCount];
          updatedColCount += 1;
        }
        return updatedColValue;
      });
      updatedRowCount += 1;
    }
    return updateRows;
  });
};

export interface IFormattedSnapshotResponse {
  showableTable: string[][];
  auditIdTable: string[][];
}

export const transformSnapshotResponse = (functionResponse: IFunctionResponse): IFormattedSnapshotResponse => {
  const showableTable: string[][] = [];
  const auditIdTable: string[][] = [];
  if (functionResponse.results) {
    const sortedResults = functionResponse.results.sort((prev: IFunctionResultObject, next: IFunctionResultObject) => {
      return prev.position.row - next.position.row;
    });
    sortedResults.forEach((result, index: number) => {
      // need to check for {} because an issue from OPS ticket: EFO-13611
      const dataList: (string | number | null)[][] = result.data.map((rowDataList) => {
        return rowDataList.map((colData) => {
          if (typeof colData === typeof {} && colData && Object.keys(colData).length === 0) {
            return "#N/A";
          }
          if (colData === null) {
            return "NULL";
          }
          return colData;
        });
      }) as (string | number | null)[][];

      showableTable.push(
        ...mapDataToTableFormat(dataList as string[][], index, result.position, functionResponse.callerValue as string)
      );
      if (Array.isArray(result.clickThrough)) {
        auditIdTable.push(
          ...mapDataToTableFormat(
            result.clickThrough as string[][],
            index,
            result.position,
            functionResponse.callerClickThrough
          )
        );
      }
    });
  }
  return {
    showableTable,
    auditIdTable,
  };
};

const mapDataToTableFormat = (
  data: string[][],
  index: number,
  position: {
    row: number;
    column: number;
  },
  callerValue?: string
): string[][] => {
  const mutableData: string[][] = [...data] as string[][];
  if (index === 0) {
    if (position.row === 0) {
      mutableData[0].unshift(callerValue || "");
    } else {
      mutableData.unshift(callerValue ? [callerValue] : [""]);
    }
  }
  return mutableData;
};
