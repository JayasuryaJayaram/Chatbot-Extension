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

  public onInit(): Promise<void> {
    // Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

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
            <iframe src="https://copilotstudio.microsoft.com/environments/Default-7d329492-602b-4902-8434-ce53aa47b425/bots/cr33a_copilotSp/webchat?__version__=2"
            frameborder="0" style="width: 350px; height: 500px; border-radius: 10px;" ></iframe>
            </div>
        </div>
    </div>`;

          const popoverButton =
            this._bottomPlaceholder.domElement.querySelector(
              "#popoverButton"
            ) as HTMLButtonElement;
          const popoverContent =
            this._bottomPlaceholder.domElement.querySelector(
              "#popoverContent"
            ) as HTMLElement;

          const closePopover = () => {
            popoverContent.style.display = "none";
            document.body.removeEventListener("click", closePopover);
          };

          if (popoverButton && popoverContent) {
            popoverButton.addEventListener("click", (event) => {
              event.stopPropagation(); // Prevents the click event from reaching the document body
              if (popoverContent.style.display === "none") {
                popoverContent.style.display = "block";
                document.body.addEventListener("click", closePopover);
              } else {
                popoverContent.style.display = "none";
                document.body.removeEventListener("click", closePopover);
              }
            });

            // Prevent clicks inside the popover from closing it
            popoverContent.addEventListener("click", (event) => {
              event.stopPropagation(); // Prevents the click event from reaching the document body
            });
          }
        }
      }
    }
  }

  private _onDispose(): void {
    console.log(
      "[ChatbotApplicationCustomizer._onDispose] Disposed custom bottom placeholder."
    );
  }
}
