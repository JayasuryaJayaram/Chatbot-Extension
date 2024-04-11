import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName,
} from "@microsoft/sp-application-base";
import styles from "./AppCustomizer.module.scss";

export interface IChatbotApplicationCustomizerProperties {
  footer: string;
}

export default class ChatbotApplicationCustomizer extends BaseApplicationCustomizer<IChatbotApplicationCustomizerProperties> {
  private _bottomPlaceholder: PlaceholderContent | undefined;
  private _popoverContent: HTMLElement | undefined;

  public onInit(): Promise<void> {
    this.context.placeholderProvider.changedEvent.add(
      this,
      this._renderPlaceHolders
    );

    return Promise.resolve();
  }

  private _renderPlaceHolders(): void {
    if (!this._bottomPlaceholder) {
      this._bottomPlaceholder =
        this.context.placeholderProvider.tryCreateContent(
          PlaceholderName.Bottom,
          { onDispose: this._onDispose }
        );

      if (!this._bottomPlaceholder) {
        console.error("The expected placeholder (Bottom) was not found.");
        return;
      }

      if (this.properties) {
        if (this._bottomPlaceholder.domElement) {
          this._bottomPlaceholder.domElement.innerHTML = `
            <div class="${styles.app}">
              <div class="${styles.bottom}">
                <button id="popoverButton" class="${styles.btn}">
                  <img src=${require("./OnlineSupport.png")} alt="Chatbot logo" style='width: 85%;'>
                </button>
                <div id="popoverContent" class="${
                  styles.popoverContent
                }" style="display: none;">               
                  <iframe src="https://copilotstudio.microsoft.com/environments/Default-7d329492-602b-4902-8434-ce53aa47b425/bots/cr33a_copilotSp/webchat?__version__=2" frameborder="0" style="width: 350px; height: 500px; border-radius: 10px; box-shadow: 2px 1px 5px 2px #978f8f"></iframe>
                </div>
              </div>
            </div>`;

          const popoverButton =
            this._bottomPlaceholder.domElement.querySelector(
              "#popoverButton"
            ) as HTMLButtonElement;
          this._popoverContent =
            this._bottomPlaceholder.domElement.querySelector(
              "#popoverContent"
            ) as HTMLElement;

          if (popoverButton && this._popoverContent) {
            popoverButton.addEventListener("click", this._togglePopover);

            this._popoverContent.addEventListener("click", (event) => {
              event.stopPropagation(); // Prevents the click event from reaching the document body
            });
          }
        }
      }
    }
  }

  private _togglePopover = (event: MouseEvent): void => {
    event.stopPropagation(); // Prevents the click event from reaching the document body

    if (this._popoverContent) {
      this._popoverContent.style.display =
        this._popoverContent.style.display === "none" ? "block" : "none";

      if (this._popoverContent.style.display === "block") {
        document.body.addEventListener(
          "click",
          this._closePopoverOnOutsideClick
        );
        window.addEventListener("scroll", this._closePopoverOnScroll, true); // Add scroll event listener
      } else {
        document.body.removeEventListener(
          "click",
          this._closePopoverOnOutsideClick
        );
        window.removeEventListener("scroll", this._closePopoverOnScroll, true); // Remove scroll event listener
      }
    }
  };

  private _closePopoverOnOutsideClick = (): void => {
    if (this._popoverContent) {
      this._popoverContent.style.display = "none";
      document.body.removeEventListener(
        "click",
        this._closePopoverOnOutsideClick
      );
      window.removeEventListener("scroll", this._closePopoverOnScroll, true);
    }
  };

  private _closePopoverOnScroll = (): void => {
    if (this._popoverContent) {
      this._popoverContent.style.display = "none";
      window.removeEventListener("scroll", this._closePopoverOnScroll, true);
    }
  };

  private _onDispose(): void {
    console.log(
      "[ChatbotApplicationCustomizer._onDispose] Disposed custom bottom placeholder."
    );
  }
}
