#include <WiFi.h>
#include <HTTPClient.h>
#include <math.h>

const char* WIFI_SSID = "SUA_REDE_WIFI";
const char* WIFI_SENHA = "SUA_SENHA_WIFI";
const char* API_URL = "http://192.168.1.100:5000/api/energia/leitura";
const char* DISPOSITIVO = "Quadro principal";

const int PINO_TENSAO = 34;
const int PINO_CORRENTE = 35;
const float TENSAO_REDE_V = 220.0;
const float TARIFA_KWH = 0.95;
const float SENSIBILIDADE_CORRENTE = 0.100;
const float OFFSET_CORRENTE_A = 0.0;
const unsigned long INTERVALO_MS = 60000;

unsigned long ultimaLeitura = 0;

void conectarWifi() {
  WiFi.mode(WIFI_STA);
  WiFi.begin(WIFI_SSID, WIFI_SENHA);
  Serial.print("Conectando ao Wi-Fi");
  while (WiFi.status() != WL_CONNECTED) {
    delay(500);
    Serial.print(".");
  }
  Serial.println();
  Serial.print("ESP32 conectado em: ");
  Serial.println(WiFi.localIP());
}

float lerCorrenteRms() {
  const int amostras = 200;
  double soma = 0;
  for (int i = 0; i < amostras; i++) {
    float tensaoSensor = analogReadMilliVolts(PINO_CORRENTE) / 1000.0;
    soma += tensaoSensor * tensaoSensor;
    delayMicroseconds(200);
  }
  float tensaoRms = sqrt(soma / amostras);
  float corrente = (tensaoRms / SENSIBILIDADE_CORRENTE) - OFFSET_CORRENTE_A;
  return max(0.0f, corrente);
}

void enviarLeitura(float corrente, float potencia, float energia, float custo) {
  if (WiFi.status() != WL_CONNECTED) conectarWifi();

  HTTPClient http;
  http.begin(API_URL);
  http.addHeader("Content-Type", "application/json");
  String json = "{";
  json += "\"dispositivo\":\"" + String(DISPOSITIVO) + "\",";
  json += "\"energia_kwh\":" + String(energia, 5) + ",";
  json += "\"custo_brl\":" + String(custo, 2) + ",";
  json += "\"potencia_w\":" + String(potencia, 2) + ",";
  json += "\"tensao_v\":" + String(TENSAO_REDE_V, 2) + ",";
  json += "\"corrente_a\":" + String(corrente, 3);
  json += "}";

  int resposta = http.POST(json);
  Serial.printf("Envio de energia: HTTP %d\n", resposta);
  if (resposta > 0) Serial.println(http.getString());
  http.end();
}

void setup() {
  Serial.begin(115200);
  analogReadResolution(12);
  conectarWifi();
  ultimaLeitura = millis();
}

void loop() {
  unsigned long agora = millis();
  if (agora - ultimaLeitura < INTERVALO_MS) {
    delay(100);
    return;
  }

  float horas = (agora - ultimaLeitura) / 3600000.0;
  ultimaLeitura = agora;
  float corrente = lerCorrenteRms();
  float potencia = TENSAO_REDE_V * corrente;
  float energia = potencia * horas / 1000.0;
  float custo = energia * TARIFA_KWH;
  enviarLeitura(corrente, potencia, energia, custo);
}
