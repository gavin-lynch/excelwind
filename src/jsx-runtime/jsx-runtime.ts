/// <reference path="../jsx-types.d.ts" />
function jsx(type: any, props: any) {
  // The 'key' argument is for React compatibility, but we can ignore it here.
  // The custom components are functions that return the element object.
  if (typeof type === 'function') {
    return type(props);
  }
  // For potential intrinsic elements (though not used in this project).
  return { type, props };
}

// jsxs is the same as jsx for this simple implementation.
// It's an optimization for multiple children in React.
export { jsx, jsx as jsxs };
export const Fragment = ({ children }: { children: any }) => children;
