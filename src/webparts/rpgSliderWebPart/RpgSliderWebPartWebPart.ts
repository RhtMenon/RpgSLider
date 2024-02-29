import { BaseClientSideWebPart, IPropertyPaneConfiguration, PropertyPaneDropdown } from '@microsoft/sp-webpart-base';
import styles from './RpgSliderWebPartWebPart.module.scss';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export interface IRpgSliderWebpartWebPartProps {
  listName: string;
}

export interface ILogoItem {
  Title: string;
  LinkUrl:{
    Url: string;
  };
  LogoUrl: string;
}

export default class RpgSliderWebpartWebPart extends BaseClientSideWebPart<IRpgSliderWebpartWebPartProps> {
  private logoItems: ILogoItem[] = [];
  private availableLists: { key: string; text: string }[] = [];
  private sliderContainer: HTMLDivElement | null = null;
  private logoIndex: number = 0;

  protected onInit(): Promise<void> {
    // Load available lists and logo data before rendering
    return Promise.all([this.loadAvailableLists(), this.loadLogoData()]).then(() => {
      this.render();
    });
  }

  public render(): void {
    // Check if logos are loaded before rendering
    if (this.logoItems.length === 0) {
      return;
    }

    // Check if the slider container has already been created
    if (!this.sliderContainer) {
      const container = document.createElement('div');
      container.className = styles.sliderParentDiv;

      this.sliderContainer = document.createElement('div');
      this.sliderContainer.id = 'client-slider';
      this.sliderContainer.className = styles.clientSlider;

      // Add arrow buttons for manual control
      const prevButton = document.createElement('button');
      prevButton.innerHTML = '<'; // You can use an arrow icon or an image here
      prevButton.className = styles.arrowleft;
      prevButton.onclick = () => this.slide(-1);

      const nextButton = document.createElement('button');
      nextButton.innerHTML = '>'; // You can use an arrow icon or an image here
      nextButton.className = styles.arrowright;
      nextButton.onclick = () => this.slide(1);

      container.appendChild(prevButton);
      container.appendChild(this.sliderContainer);
      container.appendChild(nextButton);

      this.domElement.appendChild(container);

      // Duplicate logos for continuous loop
      this.renderLogos(this.sliderContainer, this.logoItems.concat(this.logoItems));
    }
  }

  private slide(direction: number): void {
    const logoWidth = this.getLogoWidth();
    const containerWidth = this.sliderContainer?.offsetWidth || 0;
    const scrollLeft = this.sliderContainer?.scrollLeft || 0;
    const scrollWidth = this.sliderContainer?.scrollWidth || 0;

    // Calculate the next logo index based on the direction
    this.logoIndex = (this.logoIndex + direction + this.logoItems.length) % this.logoItems.length;

    if (this.sliderContainer) {
      // Calculate the new position without using 'behavior: smooth'
      const newPosition = scrollLeft + direction * logoWidth;

      // Check if reached the end, then instantly reset to the beginning or end
      if (direction > 0 && newPosition + containerWidth >= scrollWidth) {
        this.sliderContainer.scrollLeft = 0;
      } else if (direction < 0 && newPosition <= 0) {
        this.sliderContainer.scrollLeft = scrollWidth - containerWidth;
      } else {
        // Otherwise, set the new position with smooth scrolling
        this.sliderContainer.scrollTo({
          left: newPosition,
          behavior: 'smooth',
        });
      }
    }
  }

  private getLogoWidth(): number {
    const logos = this.sliderContainer?.getElementsByTagName('img');
    const firstLogo = logos?.[0];

    if (firstLogo) {
      return firstLogo.offsetWidth + 5; // Adjust as needed
    }

    return 0;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: 'Configure Logo Slider',
          },
          groups: [
            {
              groupName: 'Settings',
              groupFields: [
                PropertyPaneDropdown('listName', {
                  label: 'Select Logo List',
                  options: this.availableLists,
                }),
              ],
            },
          ],
        },
      ],
    };
  }

  private loadAvailableLists(): Promise<void> {
    const endpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists?$select=Title`;

    return this.context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => response.json())
      .then((data: { value: { Title: string }[] }) => {
        this.availableLists = data.value.map((list) => ({ key: list.Title, text: list.Title }));
        this.context.propertyPane.refresh();
      })
      .catch((error) => {
        console.error('Error fetching available lists', error);
      });
  }

  private loadLogoData(): Promise<void> {
    const { listName } = this.properties;

    if (listName) {
      const endpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items?$select=Title,LogoUrl,LinkUrl`;

      return this.context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1)
        .then((response: SPHttpClientResponse) => response.json())
        .then((data: { value: ILogoItem[] }) => {
          console.log('API response:', data); // Log the entire API response
          this.logoItems = data.value;
        })
        .catch((error) => {
          console.error('Error loading logo data', error);
        });
    }

    return Promise.resolve();
  }

  private renderLogos(container: HTMLElement, logos: ILogoItem[]): void {
    logos.forEach((logo) => {
      const img = document.createElement('img');
      img.src = logo.LogoUrl;
      img.alt = logo.Title;
      img.onclick = () => window.open(logo.LinkUrl.Url, '_self');
      container.appendChild(img);
    });
  }
}