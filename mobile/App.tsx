import { StatusBar } from "expo-status-bar";
import { useEffect, useRef, useState } from "react";
import { ActivityIndicator, BackHandler, StyleSheet, Text, View } from "react-native";
import { WebView, type WebViewNavigation } from "react-native-webview";

// The Android application deliberately contains no parallel native workflow.
// The website is the single source of truth for upload, rendering, placement,
// exceptions, search/replace and export behavior. The WebView renders the
// GitHub Pages origin (https://hautlys.github.io/assina-lote-delivery/), which
// is the same static build deployed by .github/workflows/pages.yml.
const WEBSITE_URL = "https://hautlys.github.io/assina-lote-delivery/";

export default function App() {
  const webViewRef = useRef<WebView>(null);
  const [canGoBack, setCanGoBack] = useState(false);
  const [failed, setFailed] = useState(false);

  useEffect(() => {
    const subscription = BackHandler.addEventListener("hardwareBackPress", () => {
      if (!canGoBack) return false;
      webViewRef.current?.goBack();
      return true;
    });
    return () => subscription.remove();
  }, [canGoBack]);

  if (failed) {
    return (
      <View style={styles.error}>
        <Text style={styles.errorTitle}>Assina Lote</Text>
        <Text style={styles.errorText}>Não foi possível carregar o site. Verifique sua conexão e abra o aplicativo novamente.</Text>
      </View>
    );
  }

  return (
    <View style={styles.root}>
      <StatusBar style="dark" />
      <WebView
        ref={webViewRef}
        source={{ uri: WEBSITE_URL }}
        style={styles.webView}
        javaScriptEnabled
        domStorageEnabled
        sharedCookiesEnabled
        thirdPartyCookiesEnabled
        allowsInlineMediaPlayback
        mediaPlaybackRequiresUserAction={false}
        allowFileAccess
        allowUniversalAccessFromFileURLs
        setSupportMultipleWindows={false}
        startInLoadingState
        onNavigationStateChange={(state: WebViewNavigation) => setCanGoBack(state.canGoBack)}
        onError={() => setFailed(true)}
        onHttpError={(event) => { if (event.nativeEvent.statusCode >= 500) setFailed(true); }}
        renderLoading={() => (
          <View style={styles.loading}>
            <ActivityIndicator size="large" color="#173f4b" />
            <Text style={styles.loadingText}>Carregando Assina Lote…</Text>
          </View>
        )}
      />
    </View>
  );
}

const styles = StyleSheet.create({
  root: { flex: 1, backgroundColor: "#f5f6f2" },
  webView: { flex: 1, backgroundColor: "#f5f6f2" },
  loading: { ...StyleSheet.absoluteFill, alignItems: "center", justifyContent: "center", gap: 12, backgroundColor: "#f5f6f2" },
  loadingText: { color: "#173f4b", fontSize: 15, fontWeight: "600" },
  error: { flex: 1, padding: 28, alignItems: "center", justifyContent: "center", backgroundColor: "#f5f6f2" },
  errorTitle: { color: "#173f4b", fontSize: 24, fontWeight: "800", marginBottom: 12 },
  errorText: { color: "#647873", fontSize: 15, lineHeight: 22, textAlign: "center" },
});
