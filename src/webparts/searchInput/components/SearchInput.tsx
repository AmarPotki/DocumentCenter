import * as React from 'react';
import styles from './SearchInput.module.scss';
import { ISearchInputProps } from './ISearchInputProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { IUtilities, Utilities } from "../../../common";
import {
  RecognizerConfig, SpeechConfig,
  Recognizer, Context, OS, Device, CognitiveSubscriptionKeyAuthentication,
  RecognitionMode, SpeechResultFormat, SpeechRecognitionEvent, RecognitionTriggeredEvent,
  ListeningStartedEvent, RecognitionStartedEvent, SpeechStartDetectedEvent, SpeechHypothesisEvent,
  SpeechEndDetectedEvent, SpeechSimplePhraseEvent, SpeechDetailedPhraseEvent, RecognitionEndedEvent, IDetailedSpeechPhrase
} from "../../../bingSpeechApi/sdk/speech/Exports";
import { CreateRecognizer } from "../../../bingSpeechApi/sdk/speech.browser/Exports";
import { BingSpeech } from "../../../bingSpeechApi/textToSpeach/BingSpeech";
export interface ISearchInputState {
  value: string;
  Icon: string;
  isSearchIcon: boolean;
  microphonStyle: string;
  key: string;
}
export class RecognizedPhraseResult {
  public RecognitionStatus: string;
  public DisplayText: string;
}
export default class SearchInput extends React.Component<ISearchInputProps, ISearchInputState> {
  private utilities: IUtilities;
  /**
   *
   */
  constructor(props: ISearchInputProps) {
    super(props);
    let value: string = '';
    let url: any = new URL(window.location.href);
    if (url.searchParams.has(this.props.keyword)) {
      value = url.searchParams.get(this.props.keyword);
    }
    this.utilities = new Utilities();
    let micIcon = this.utilities.getIcon("Microphon", false);
    this.state = {
      value: value, Icon: micIcon, isSearchIcon: false,
      microphonStyle: '', key: "8f4b2dc63d7e4430a0f7ca2d9711f716"
    };
  }
  public render(): React.ReactElement<ISearchInputProps> {
    return (
      <div className={styles.searchInput} >
        <div className={styles.container}>
          <div className="ms-Grid">
            <div className={"ms-Grid-row ms-fontColor-white " + styles.row}>
              <div className={"ms-Grid-col ms-lg12 ms-md12 ms-sm12 " + styles.column}>
                <div className="search-filter-block">
                  <div className={styles.title}> {this.props.title}</div>
                  <div className={styles.embedAfield}>
                    <input type="text" onChange={(e) => this.onChanged(e)}
                      value={this.state.value}
                      onKeyPress={(e) => this.onKeyPress(e)} placeholder={this.props.placeholder} />
                    {this.state.isSearchIcon ?
                      (<a onClick={() => this.updateQueryString()}>
                        <img src={this.utilities.getIcon("Magnifier", false)} /></a>) : ""}
                    <a onClick={() => this.onStart()} className={this.state.microphonStyle}><img src={this.state.Icon} /></a>

                  </div>

                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }
  private onChanged(event: any): void {
    let controlId: any = event.target.id;
    let value: string = event.target.value;
    if (value.length > 0) {
      this.setState({
        isSearchIcon: true, value: value,
        microphonStyle: styles.atagWithSearchIcon
      });

    }
    else {
      this.setState({
        value: value,
        Icon: this.utilities.getIcon("Microphon", false)
      });
    }
  }

  private onKeyPress(event: any): void {
    if (event.charCode === 13) {
      event.preventDefault();
      event.stopPropagation();
      this.updateQueryString();
    }
  }
  private recognizer: Recognizer;
  public setup() {
    let recognitionMode: RecognitionMode = RecognitionMode.Interactive;
    let language: string = 'en-GB';
    let format: SpeechResultFormat = SpeechResultFormat.Simple;
    let recognizerConfig = new RecognizerConfig(
      new SpeechConfig(
        new Context(
          new OS(navigator.userAgent, "Browser", null),
          new Device("SpeechSample", "SpeechSample", "1.0.00000"))),
      recognitionMode, // SDK.RecognitionMode.Interactive  (Options - Interactive/Conversation/Dictation)
      language, // Supported laguages are specific to each recognition mode. Refer to docs.
      format); // SDK.SpeechResultFormat.Simple (Options - Simple/Detailed)

    // Alternatively use SDK.CognitiveTokenAuthentication(fetchCallback, fetchOnExpiryCallback) for token auth
    let authentication = new CognitiveSubscriptionKeyAuthentication(this.state.key);
    this.recognizer = CreateRecognizer(recognizerConfig, authentication);

  }
  public onStart(): void {
    this.setState({ Icon: this.utilities.getIcon("MicrophonRed", false) });
    this.setup();
    this.recognizer.Recognize((event: SpeechRecognitionEvent) => {
      if (event instanceof RecognitionTriggeredEvent) {
        this.UpdateStatus("Initializing test");
      } else if (event instanceof ListeningStartedEvent) {
        this.UpdateStatus("Listening Started");
      } else if (event instanceof RecognitionStartedEvent) {
        this.UpdateStatus("Listening_Recognizing");
      } else if (event instanceof SpeechStartDetectedEvent) {
        this.UpdateStatus("Listening_DetectedSpeech_Recognizing");
      } else if (event instanceof SpeechHypothesisEvent) {
        this.UpdateRecognizedHypothesis(event.Result.Text);
      } else if (event instanceof SpeechEndDetectedEvent) {
        this.OnSpeechEndDetected();
      } else if (event instanceof SpeechSimplePhraseEvent) {
        this.UpdateRecognizedPhrase(JSON.stringify(event.Result, null, 3));
      } else if (event instanceof SpeechDetailedPhraseEvent) {
        this.UpdateRecognizedPhrase(JSON.stringify(event.Result, null, 3));
        // this.UpdateRecognizedPhrase(event.Result);
      } else if (event instanceof RecognitionEndedEvent) {
        this.OnComplete();
        this.UpdateStatus("Idle");
      }
    }
    )
      .On(() => {
        // The request succeeded. Nothing to do here.
        console.log(`speeCh`);

      },
      (error: any) => {
        //  console.error(error);
      });
  }
  private UpdateStatus(value: string) {
    // console.log(`Status: ${value}`);
  }

  private UpdateRecognizedHypothesis(text: string) {
    // console.log(`Hypothesis: ${text}`);
    this.setState({
      value: text
    });
  }

  private OnSpeechEndDetected() {
    // console.log(`SpeechEndDetected`);
  }
  private UpdateRecognizedPhrase(phrase: string) {
    let result = JSON.parse(phrase) as RecognizedPhraseResult;
    // console.log(result);

    if (result.RecognitionStatus === "InitialSilenceTimeout") {
      this.setState({
        Icon: this.utilities.getIcon("Microphon", false)
      });
    } else {
      //let text = result.DisplayText.endsWith(".") ? result.DisplayText.slice(0, -1) : result.DisplayText;
      let text = result.DisplayText ? result.DisplayText.slice(0, -1) : result.DisplayText;
      this.setState({
        isSearchIcon: true,
        value: text,
        Icon: this.utilities.getIcon("Microphon", false),
        microphonStyle: styles.atagWithSearchIcon
      });
      this.updateQueryString();
      console.log("speach");
      let bingClientTTS = new BingSpeech.TTSClient(this.state.key);
      bingClientTTS.synthesize(`Here are the search results for ${text}`, BingSpeech.SupportedLocales.enGB_Female);
    }

  }
  private OnComplete() {

  }

  private updateQueryString() {
    let url: any = new URL(window.location.href);
    if (url.searchParams.has(this.props.keyword)) {
      url.searchParams.delete(this.props.keyword);
    }
    url.searchParams.set(this.props.keyword, this.state.value);
    if (history.pushState) {
      window.history.pushState({ path: url.href }, '', url.href);
    }
  }
}
