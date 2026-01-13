// In development, the JSX transform uses jsxDEV for more detailed error messages.
function jsxDEV(type: any, props: any) {
  if (typeof type === 'function') {
    return type(props);
  }
  return { type, props };
}

export { jsxDEV };
