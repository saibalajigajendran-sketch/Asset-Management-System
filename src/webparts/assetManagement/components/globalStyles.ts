export const fullScreenBodyClass = 'assetManagementFullScreen';

const fullScreenStyleId = 'asset-management-full-screen-styles';

export const ensureFullScreenStyles = (): void => {
  if (document.getElementById(fullScreenStyleId)) {
    return;
  }

  const style = document.createElement('style');
  style.id = fullScreenStyleId;
  style.textContent = `
    body.${fullScreenBodyClass} {
      height: 100vh !important;
      margin: 0 !important;
      overflow: hidden !important;
    }
    body.${fullScreenBodyClass} #SuiteNavWrapper,
    body.${fullScreenBodyClass} #sp-appBar,
    body.${fullScreenBodyClass} #spLeftNav,
    body.${fullScreenBodyClass} #spSiteHeader,
    body.${fullScreenBodyClass} .ms-HubNav,
    body.${fullScreenBodyClass} [data-automation-id="SiteHeader"],
    body.${fullScreenBodyClass} [data-automation-id="CommandBar"],
    body.${fullScreenBodyClass} [data-automation-id="pageCommandBar"],
    body.${fullScreenBodyClass} [data-automation-id="pageFooter"],
    body.${fullScreenBodyClass} .ms-SocialBar,
    body.${fullScreenBodyClass} [data-sp-feature-tag="Comments"] {
      display: none !important;
    }
    body.${fullScreenBodyClass} #spPageCanvasContent,
    body.${fullScreenBodyClass} .CanvasZoneContainer,
    body.${fullScreenBodyClass} .CanvasZone,
    body.${fullScreenBodyClass} .ControlZone,
    body.${fullScreenBodyClass} [class*="largeContainer"] {
      width: 100% !important;
      max-width: none !important;
      margin-left: 0 !important;
      margin-right: 0 !important;
      padding-left: 0 !important;
      padding-right: 0 !important;
    }
  `;
  document.head.appendChild(style);
};