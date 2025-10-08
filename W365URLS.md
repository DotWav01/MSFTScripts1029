# Windows 365 & Azure Virtual Desktop Firewall Whitelist
## Quick Reference Guide for Network Engineers

---

## 1. WINDOWS 365 CORE SERVICES

### Primary Infrastructure Endpoints
| FQDN/Address | Protocol | Port(s) | Purpose |
|--------------|----------|---------|---------|
| *.infra.windows365.microsoft.com | TCP | 443 | Core W365 infrastructure |
| *.cmdagent.trafficmanager.net | TCP | 443 | Command agent traffic |

### Authentication Endpoints
| FQDN/Address | Protocol | Port(s) | Purpose |
|--------------|----------|---------|---------|
| login.microsoftonline.com | TCP | 443 | Microsoft Online Services authentication |
| login.live.com | TCP | 443 | Microsoft account authentication |
| enterpriseregistration.windows.net | TCP | 80, 443 | Device registration |

### Azure IoT Device Provisioning
| FQDN/Address | Protocol | Port(s) | Purpose |
|--------------|----------|---------|---------|
| global.azure-devices-provisioning.net | TCP | 443, 5671 | Global provisioning service |
| hm-iot-in-prod-prap01.azure-devices.net | TCP | 443, 5671 | IoT Hub - Asia Pacific |
| hm-iot-in-prod-prau01.azure-devices.net | TCP | 443, 5671 | IoT Hub - Australia |
| hm-iot-in-prod-preu01.azure-devices.net | TCP | 443, 5671 | IoT Hub - Europe |
| hm-iot-in-prod-prna01.azure-devices.net | TCP | 443, 5671 | IoT Hub - North America 1 |
| hm-iot-in-prod-prna02.azure-devices.net | TCP | 443, 5671 | IoT Hub - North America 2 |
| hm-iot-in-2-prod-preu01.azure-devices.net | TCP | 443, 5671 | IoT Hub Gen2 - Europe |
| hm-iot-in-2-prod-prna01.azure-devices.net | TCP | 443, 5671 | IoT Hub Gen2 - North America |
| hm-iot-in-3-prod-preu01.azure-devices.net | TCP | 443, 5671 | IoT Hub Gen3 - Europe |
| hm-iot-in-3-prod-prna01.azure-devices.net | TCP | 443, 5671 | IoT Hub Gen3 - North America |
| hm-iot-in-4-prod-prna01.azure-devices.net | TCP | 443, 5671 | IoT Hub Gen4 - North America |

---

## 2. AZURE VIRTUAL DESKTOP (REQUIRED)

### Core AVD Services
| FQDN/Address | Protocol | Port(s) | Purpose | Service Tag |
|--------------|----------|---------|---------|-------------|
| login.microsoftonline.com | TCP | 443 | Authentication | AzureActiveDirectory |
| *.wvd.microsoft.com | TCP | 443 | AVD service traffic | WindowsVirtualDesktop |
| catalogartifact.azureedge.net | TCP | 443 | Azure Marketplace | AzureFrontDoor.Frontend |
| *.prod.warm.ingest.monitor.core.windows.net | TCP | 443 | Diagnostics | AzureMonitor |
| gcs.prod.monitoring.core.windows.net | TCP | 443 | Agent monitoring | AzureMonitor |
| azkms.core.windows.net | TCP | 1688 | Windows activation | Internet |
| mrsglobalsteus2prod.blob.core.windows.net | TCP | 443 | Agent/SXS stack updates | Storage |
| wvdportalstorageblob.blob.core.windows.net | TCP | 443 | Azure portal support | AzureCloud |
| 169.254.169.254 | TCP | 80 | Azure metadata service | N/A |
| 168.63.129.16 | TCP | 80 | Session host health monitoring | N/A |

### Certificate Services
| FQDN/Address | Protocol | Port(s) | Purpose | Service Tag |
|--------------|----------|---------|---------|-------------|
| oneocsp.microsoft.com | TCP | 80 | Certificate validation | AzureFrontDoor.FirstParty |
| www.microsoft.com | TCP | 80 | Certificate services | N/A |
| ctldl.windowsupdate.com | TCP | 80 | Certificate trust lists | N/A |

### Additional Service Traffic
| FQDN/Address | Protocol | Port(s) | Purpose | Service Tag |
|--------------|----------|---------|---------|-------------|
| aka.ms | TCP | 443 | URL shortener (Azure Local) | N/A |
| *.service.windows.cloud.microsoft | TCP | 443 | Service traffic | WindowsVirtualDesktop |
| *.windows.cloud.microsoft | TCP | 443 | Service traffic | N/A |
| *.windows.static.microsoft | TCP | 443 | Service traffic | N/A |

---

## 3. AZURE VIRTUAL DESKTOP (OPTIONAL)

### Authentication & Microsoft 365
| FQDN/Address | Protocol | Port(s) | Purpose | Service Tag |
|--------------|----------|---------|---------|-------------|
| login.windows.net | TCP | 443 | Microsoft 365 sign-in | AzureActiveDirectory |
| *.events.data.microsoft.com | TCP | 443 | Telemetry | N/A |

### Network & Updates
| FQDN/Address | Protocol | Port(s) | Purpose | Service Tag |
|--------------|----------|---------|---------|-------------|
| www.msftconnecttest.com | TCP | 80 | Internet connectivity check | N/A |
| *.prod.do.dsp.mp.microsoft.com | TCP | 443 | Windows Update | N/A |
| *.sfx.ms | TCP | 443 | OneDrive client updates | N/A |
| *.digicert.com | TCP | 80 | Certificate revocation | N/A |

### Azure DNS
| FQDN/Address | Protocol | Port(s) | Purpose | Service Tag |
|--------------|----------|---------|---------|-------------|
| *.azure-dns.com | TCP | 443 | Azure DNS resolution | N/A |
| *.azure-dns.net | TCP | 443 | Azure DNS resolution | N/A |

### Diagnostics
| FQDN/Address | Protocol | Port(s) | Purpose | Service Tag |
|--------------|----------|---------|---------|-------------|
| *.eh.servicebus.windows.net | TCP | 443 | Diagnostic settings | EventHub |

---

## 4. INTUNE MANAGEMENT

### Core Intune Services
| FQDN/Address | Protocol | Port(s) | Purpose |
|--------------|----------|---------|---------|
| *.manage.microsoft.com | TCP | 80, 443 | Intune client/host service |
| manage.microsoft.com | TCP | 80, 443 | Intune management |
| EnterpriseEnrollment.manage.microsoft.com | TCP | 80, 443 | Device enrollment |

### Intune IP Ranges (if domain-based rules not supported)
104.46.162.96/27, 13.67.13.176/28, 13.67.15.128/27, 13.69.231.128/28,
13.69.67.224/28, 13.70.78.128/28, 13.70.79.128/27, 13.74.111.192/27,
13.77.53.176/28, 13.86.221.176/28, 13.89.174.240/28, 13.89.175.192/28,
20.189.229.0/25, 20.191.167.0/25, 20.37.153.0/24, 20.37.192.128/25,
20.38.81.0/24, 20.41.1.0/24, 20.42.1.0/24, 20.42.130.0/24,
20.42.224.128/25, 20.43.129.0/24, 20.44.19.224/27, 40.119.8.128/25,
40.67.121.224/27, 40.70.151.32/28, 40.71.14.96/28, 40.74.25.0/24,
40.78.245.240/28, 40.78.247.128/27, 40.79.197.64/27, 40.79.197.96/28,
40.80.180.208/28, 40.80.180.224/27, 40.80.184.128/25, 40.82.248.224/28,
40.82.249.128/25, 52.150.137.0/25, 52.162.111.96/28, 52.168.116.128/27,
52.182.141.192/27, 52.236.189.96/27, 52.240.244.160/27, 20.204.193.12/30,
20.204.193.10/31, 20.192.174.216/29, 20.192.159.40/29, 104.208.197.64/27,
172.160.217.160/27, 172.201.237.160/27, 172.202.86.192/27, 172.205.63.0/25,
172.212.214.0/25, 172.215.131.0/27, 20.168.189.128/27, 20.199.207.192/28,
20.204.194.128/31, 20.208.149.192/27, 20.208.157.128/27, 20.214.131.176/29,
20.91.147.72/29, 4.145.74.224/27, 4.150.254.64/27, 4.154.145.224/27,
4.200.254.32/27, 4.207.244.0/27, 4.213.25.64/27, 4.213.86.128/25,
4.216.205.32/27, 4.237.143.128/25, 40.84.70.128/25, 48.218.252.128/25,
57.151.0.192/27, 57.153.235.0/25, 57.154.140.128/25, 57.154.195.0/25,
57.155.45.128/25, 68.218.134.96/27, 74.224.214.64/27, 74.242.35.0/25,
172.208.170.0/25, 74.241.231.0/25, 74.242.184.128/25

### Azure Front Door IP Ranges (Intune)
13.107.219.0/24, 13.107.227.0/24, 13.107.228.0/23, 150.171.97.0/24,
2620:1ec:40::/48, 2620:1ec:49::/48, 2620:1ec:4a::/47

### Content Delivery & Updates
| FQDN/Address | Protocol | Port(s) | Purpose |
|--------------|----------|---------|---------|
| *.do.dsp.mp.microsoft.com | TCP | 80, 443 | Delivery Optimization |
| *.dl.delivery.mp.microsoft.com | TCP | 80, 443 | Delivery Optimization |

### Win32 App Distribution
| FQDN/Address | Protocol | Port(s) | Purpose |
|--------------|----------|---------|---------|
| swda01-mscdn.manage.microsoft.com | TCP | 80, 443 | Win32 apps CDN |
| swda02-mscdn.manage.microsoft.com | TCP | 80, 443 | Win32 apps CDN |
| swdb01-mscdn.manage.microsoft.com | TCP | 80, 443 | Win32 apps CDN |
| swdb02-mscdn.manage.microsoft.com | TCP | 80, 443 | Win32 apps CDN |
| swdc01-mscdn.manage.microsoft.com | TCP | 80, 443 | Win32 apps CDN |
| swdc02-mscdn.manage.microsoft.com | TCP | 80, 443 | Win32 apps CDN |
| swdd01-mscdn.manage.microsoft.com | TCP | 80, 443 | Win32 apps CDN |
| swdd02-mscdn.manage.microsoft.com | TCP | 80, 443 | Win32 apps CDN |
| swdin01-mscdn.manage.microsoft.com | TCP | 80, 443 | Win32 apps CDN |
| swdin02-mscdn.manage.microsoft.com | TCP | 80, 443 | Win32 apps CDN |

### Microsoft Account & Consumer Services
| FQDN/Address | Protocol | Port(s) | Purpose |
|--------------|----------|---------|---------|
| account.live.com | TCP | 443 | Microsoft account services |
| login.live.com | TCP | 443 | Microsoft account authentication |

### Feature & Configuration Services
| FQDN/Address | Protocol | Port(s) | Purpose |
|--------------|----------|---------|---------|
| go.microsoft.com | TCP | 80, 443 | Endpoint discovery |
| config.edge.skype.com | TCP | 443 | Feature deployment |
| ecs.office.com | TCP | 443 | Office configuration |
| fd.api.orgmsg.microsoft.com | TCP | 443 | Organizational messages |
| ris.prod.api.personalization.ideas.microsoft.com | TCP | 443 | Personalization services |

---

## 5. AUTHENTICATION DEPENDENCIES

### Microsoft Entra ID (Azure AD)
| FQDN/Address | Protocol | Port(s) | Purpose |
|--------------|----------|---------|---------|
| login.microsoftonline.com | TCP | 80, 443 | Primary authentication |
| graph.windows.net | TCP | 80, 443 | Graph API |

### Office 365 Configuration
| FQDN/Address | Protocol | Port(s) | Purpose |
|--------------|----------|---------|---------|
| *.officeconfig.msocdn.com | TCP | 443 | Office customization |
| config.office.com | TCP | 443 | Office policy management |

### Device Registration
| FQDN/Address | Protocol | Port(s) | Purpose |
|--------------|----------|---------|---------|
| enterpriseregistration.windows.net | TCP | 80, 443 | Device registration |
| certauth.enterpriseregistration.windows.net | TCP | 80, 443 | Certificate-based auth |

---

## IMPLEMENTATION NOTES

Default Port: Unless otherwise specified, all endpoints use TCP port 443.

Wildcard Support: If your firewall supports wildcards in domain rules, use the wildcard entries as provided. If not, refer to the IP ranges section for Intune services.

Service Tags: Where indicated, use Azure service tags for more dynamic IP management instead of static IP ranges.

Regional Considerations: For AVD agent traffic, ensure you whitelist region-specific FQDNs based on where your session hosts are deployed. Check Event ID 3701 on session hosts for region-specific endpoints.

Priority: Focus on sections 1, 2, and 4 first for core functionality. Section 3 (AVD Optional) can be added as needed based on features used.
