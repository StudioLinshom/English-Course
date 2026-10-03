package io.github.studiolinshom.englishwords;

import android.app.Activity;
import android.content.Context;
import android.content.Intent;
import android.media.AudioManager;
import android.net.Uri;
import android.os.Bundle;
import android.speech.tts.TextToSpeech;
import android.webkit.JavascriptInterface;
import android.webkit.WebChromeClient;
import android.webkit.WebResourceError;
import android.webkit.WebResourceRequest;
import android.webkit.WebSettings;
import android.webkit.WebView;
import android.webkit.WebViewClient;
import android.widget.Toast;
import java.util.Locale;

public class MainActivity extends Activity {
    static final String APP_URL = "https://studiolinshom.github.io/English-Course/";
    static final String APP_HOST = "studiolinshom.github.io";

    WebView web;
    TextToSpeech tts;
    boolean ttsReady, ttsFailed, warnedVolume;
    String pendingText;

    @Override
    protected void onCreate(Bundle state) {
        super.onCreate(state);
        web = new WebView(this);
        setContentView(web);

        WebSettings s = web.getSettings();
        s.setJavaScriptEnabled(true);
        s.setDomStorageEnabled(true);
        s.setDatabaseEnabled(true);
        s.setMediaPlaybackRequiresUserGesture(false);

        // Enables the page's alert()/confirm() dialogs.
        web.setWebChromeClient(new WebChromeClient());
        web.setWebViewClient(new WebViewClient() {
            @Override
            public boolean shouldOverrideUrlLoading(WebView v, String url) {
                Uri u = Uri.parse(url);
                if (APP_HOST.equals(u.getHost())) return false;
                try { startActivity(new Intent(Intent.ACTION_VIEW, u)); } catch (Exception e) { }
                return true;
            }

            @Override
            public void onReceivedError(WebView v, WebResourceRequest req, WebResourceError err) {
                if (!req.isForMainFrame()) return;
                v.loadDataWithBaseURL(null,
                    "<html dir='rtl'><body style='font-family:sans-serif;text-align:center;padding:40px;background:#F7F9F4;color:#1C2B4B'>"
                    + "<h2>אין חיבור לאינטרנט</h2><p>בפעם הראשונה צריך אינטרנט כדי לטעון את האפליקציה.</p>"
                    + "<button style='font-size:18px;padding:10px 24px;border-radius:12px;border:0;background:#1F7A4D;color:#fff' "
                    + "onclick=\"location.href='" + APP_URL + "'\">נסה שוב</button></body></html>",
                    "text/html", "utf-8", null);
            }
        });

        tts = new TextToSpeech(this, new TextToSpeech.OnInitListener() {
            @Override
            public void onInit(int status) {
                if (status != TextToSpeech.SUCCESS) {
                    ttsFailed = true;
                    toast("מנוע ההקראה של הטלפון לא זמין");
                    return;
                }
                if (!setEnglish()) {
                    ttsFailed = true;
                    toast("חסר קול באנגלית בטלפון. פותח התקנה…");
                    try {
                        startActivity(new Intent(TextToSpeech.Engine.ACTION_INSTALL_TTS_DATA)
                            .addFlags(Intent.FLAG_ACTIVITY_NEW_TASK));
                    } catch (Exception e) { }
                    return;
                }
                tts.setSpeechRate(0.85f);
                ttsReady = true;
                if (pendingText != null) { say(pendingText); pendingText = null; }
            }
        });
        web.addJavascriptInterface(new Speaker(), "AndroidTTS");

        if (state != null) web.restoreState(state);
        else web.loadUrl(APP_URL);
    }

    boolean setEnglish() {
        for (Locale l : new Locale[] { Locale.US, Locale.UK, Locale.ENGLISH }) {
            if (tts.setLanguage(l) >= TextToSpeech.LANG_AVAILABLE) return true;
        }
        return false;
    }

    void say(String text) {
        AudioManager am = (AudioManager) getSystemService(Context.AUDIO_SERVICE);
        if (!warnedVolume && am != null && am.getStreamVolume(AudioManager.STREAM_MUSIC) == 0) {
            warnedVolume = true;
            toast("עוצמת המדיה על 0. הגבירו כדי לשמוע");
        }
        tts.speak(text, TextToSpeech.QUEUE_FLUSH, null, "w");
    }

    void toast(final String msg) {
        runOnUiThread(new Runnable() {
            @Override
            public void run() { Toast.makeText(MainActivity.this, msg, Toast.LENGTH_LONG).show(); }
        });
    }

    class Speaker {
        @JavascriptInterface
        public void speak(String text) {
            if (ttsReady) say(text);
            else if (ttsFailed) toast("ההקראה לא זמינה בטלפון הזה");
            else pendingText = text;
        }
    }

    @Override
    protected void onSaveInstanceState(Bundle out) {
        super.onSaveInstanceState(out);
        web.saveState(out);
    }

    @Override
    public void onBackPressed() {
        if (web.canGoBack()) web.goBack();
        else super.onBackPressed();
    }

    @Override
    protected void onDestroy() {
        if (tts != null) tts.shutdown();
        web.destroy();
        super.onDestroy();
    }
}
