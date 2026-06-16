import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import fs from 'fs'
import path from 'path'
import { exec } from 'child_process'
import { fileURLToPath } from 'url'

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// https://vite.dev/config/
export default defineConfig({
  plugins: [
    react(),
    {
      name: 'apk-builder-api',
      configureServer(server) {
        server.middlewares.use(async (req, res, next) => {
          if (req.url === '/api/build-apk' && req.method === 'POST') {
            let body = '';
            req.on('data', chunk => {
              body += chunk.toString();
            });
            req.on('end', async () => {
              try {
                const config = JSON.parse(body);
                
                // Write native configurations
                // 1. capacitor.config.json
                const capConfig = {
                  appId: config.packageName || 'com.kairos.cableapp',
                  appName: config.appName || 'Kairos Cable App',
                  webDir: 'dist'
                };
                fs.writeFileSync(
                  path.resolve(__dirname, 'capacitor.config.json'),
                  JSON.stringify(capConfig, null, 2)
                );

                // 2. MainActivity.java
                const mainActivityCode = `package ${config.packageName || 'com.kairos.cableapp'};

import android.os.Bundle;
import android.view.WindowManager;
import com.getcapacitor.BridgeActivity;

public class MainActivity extends BridgeActivity {
    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        // Force the app to run in fullscreen mode (hiding status bar)
        getWindow().setFlags(
            WindowManager.LayoutParams.FLAG_FULLSCREEN,
            WindowManager.LayoutParams.FLAG_FULLSCREEN
        );
    }
}`;
                const mainActivityPath = path.resolve(
                  __dirname,
                  'android/app/src/main/java/com/kairos/cableapp/MainActivity.java'
                );
                // Ensure directory exists
                fs.mkdirSync(path.dirname(mainActivityPath), { recursive: true });
                fs.writeFileSync(mainActivityPath, mainActivityCode);

                // 3. strings.xml
                const stringsXml = `<?xml version='1.0' encoding='utf-8'?>
<resources>
    <string name="app_name">${config.appName || 'Kairos Cable App'}</string>
    <string name="title_activity_main">${config.appName || 'Kairos Cable App'}</string>
    <string name="package_name">${config.packageName || 'com.kairos.cableapp'}</string>
    <string name="custom_url_scheme">${config.packageName || 'com.kairos.cableapp'}</string>
</resources>`;
                fs.writeFileSync(
                  path.resolve(__dirname, 'android/app/src/main/res/values/strings.xml'),
                  stringsXml
                );

                // 4. styles.xml
                const color = config.primaryColor || '#4f46e5';
                const stylesXml = `<?xml version="1.0" encoding="utf-8"?>
<resources>
    <!-- Base application theme. -->
    <style name="AppTheme" parent="Theme.AppCompat.Light.DarkActionBar">
        <item name="colorPrimary">${color}</item>
        <item name="colorPrimaryDark">${color}</item>
        <item name="colorAccent">${color}</item>
    </style>

    <style name="AppTheme.NoActionBar" parent="Theme.AppCompat.DayNight.NoActionBar">
        <item name="windowActionBar">false</item>
        <item name="windowNoTitle">true</item>
        <item name="android:background">@null</item>
        <item name="android:windowFullscreen">true</item>
    </style>

    <style name="AppTheme.NoActionBarLaunch" parent="Theme.SplashScreen">
        <item name="android:background">@drawable/splash</item>
        <item name="android:windowFullscreen">true</item>
    </style>
</resources>`;
                fs.writeFileSync(
                  path.resolve(__dirname, 'android/app/src/main/res/values/styles.xml'),
                  stylesXml
                );

                // Synchronously run compilation pipeline
                console.log('[APK Compiler Backend] Launching build processes...');
                
                exec('npm run build', { cwd: __dirname }, (buildErr, stdout, stderr) => {
                  if (buildErr) {
                    console.error('[APK Build] npm run build failed:', buildErr, stderr);
                    res.writeHead(500, { 'Content-Type': 'application/json' });
                    return res.end(JSON.stringify({ status: 'error', message: 'Vite build compilation failed: ' + buildErr.message }));
                  }
                  console.log('[APK Build] Vite build completed.');

                  exec('npx cap sync', { cwd: __dirname }, (syncErr, stdout, stderr) => {
                    if (syncErr) {
                      console.error('[APK Build] cap sync failed:', syncErr, stderr);
                      res.writeHead(500, { 'Content-Type': 'application/json' });
                      return res.end(JSON.stringify({ status: 'error', message: 'Capacitor sync failed: ' + syncErr.message }));
                    }
                    console.log('[APK Build] Capacitor platform sync completed.');

                    const gradlew = process.platform === 'win32' ? 'gradlew.bat' : './gradlew';
                    exec(`${gradlew} assembleDebug`, { cwd: path.resolve(__dirname, 'android') }, (gradleErr, stdout, stderr) => {
                      if (gradleErr) {
                        console.error('[APK Build] Gradle assembleDebug failed:', gradleErr, stderr);
                        res.writeHead(500, { 'Content-Type': 'application/json' });
                        return res.end(JSON.stringify({ status: 'error', message: 'Android Gradle packaging failed: ' + gradleErr.message }));
                      }
                      console.log('[APK Build] Gradle assembleDebug completed successfully.');

                      try {
                        const srcApk = path.resolve(__dirname, 'android/app/build/outputs/apk/debug/app-debug.apk');
                        const destApk = path.resolve(__dirname, 'public/app-release.apk');
                        
                        // Copy to public to make it downloadable
                        fs.copyFileSync(srcApk, destApk);
                        
                        // Copy to dist so it exists in production preview builds
                        const distDestApk = path.resolve(__dirname, 'dist/app-release.apk');
                        if (fs.existsSync(path.resolve(__dirname, 'dist'))) {
                          fs.copyFileSync(srcApk, distDestApk);
                        }
                        
                        console.log('[APK Build] Final APK copied successfully to web distribution root.');
                        res.writeHead(200, { 'Content-Type': 'application/json' });
                        res.end(JSON.stringify({ status: 'success', message: 'Android wrapper and APK package compiled successfully.' }));
                      } catch (copyErr) {
                        console.error('[APK Build] File copy failed:', copyErr);
                        res.writeHead(500, { 'Content-Type': 'application/json' });
                        res.end(JSON.stringify({ status: 'error', message: 'Failed to deploy compiled APK target: ' + copyErr.message }));
                      }
                    });
                  });
                });

              } catch (e) {
                console.error('[APK Build] API error:', e);
                res.writeHead(500, { 'Content-Type': 'application/json' });
                res.end(JSON.stringify({ status: 'error', message: e.message }));
              }
            });
          } else {
            next();
          }
        });
      }
    }
  ]
})
