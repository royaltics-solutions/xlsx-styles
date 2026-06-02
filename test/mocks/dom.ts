function fnReturns<T>(returnValue: T): any {
  const f: any = (...args: any[]) => {
    f.mock.calls.push(args);
    return returnValue;
  };
  f.mock = { calls: [] as any[][] };
  f.mockClear = () => { f.mock.calls = []; };
  return f;
}

export const mockElement: Record<string, any> = {
  href: "",
  download: "",
  click: fnReturns(undefined),
};

export const mockDocument: Record<string, any> = {
  createElement: (...args: any[]) => {
    mockDocument.createElement.mock.calls.push(args);
    return mockElement;
  },
  body: {
    appendChild: fnReturns(undefined),
    removeChild: fnReturns(undefined),
  },
};
(mockDocument.createElement as any).mock = { calls: [] as any[][] };
(mockDocument.createElement as any).mockClear = () => { (mockDocument.createElement as any).mock.calls = []; };

export const mockURL: Record<string, any> = {
  createObjectURL: fnReturns("blob:mock-url"),
  revokeObjectURL: fnReturns(undefined),
};

export function setupDOMMocks() {
  // @ts-ignore
  globalThis.document = mockDocument;
  // @ts-ignore
  globalThis.URL.createObjectURL = mockURL.createObjectURL;
  // @ts-ignore
  globalThis.URL.revokeObjectURL = mockURL.revokeObjectURL;
}

export function resetMocks() {
  mockElement.href = "";
  mockElement.download = "";
  mockElement.click.mockClear();
  mockDocument.createElement.mockClear();
  mockDocument.body.appendChild.mockClear();
  mockDocument.body.removeChild.mockClear();
  mockURL.createObjectURL.mockClear();
  mockURL.revokeObjectURL.mockClear();
}
