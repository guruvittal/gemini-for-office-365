#!/usr/bin/env python3
"""
Interactive & CLI Office Add-in XML Manifest Generator
Author: Carlos Augusto, Principal Architect, Google

This script generates a valid, custom Microsoft Office Add-in XML manifest
for Gemini Enterprise (Agentspace), supporting both Track 1 (WIF) and
Track 2 (Cloud Identity / Google Workspace).
"""

import argparse
import sys
import uuid
import re

WIF_TEMPLATE = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">
  <Id>{manifest_id}</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>{provider_name}</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="{display_name}"/>
  <Description DefaultValue="{description}"/>
  <IconUrl DefaultValue="{frontend_url}/assets/icon-32.png"/>
  <HighResolutionIconUrl DefaultValue="{frontend_url}/assets/icon-64.png"/>
  <SupportUrl DefaultValue="https://cloud.google.com/vertex-ai"/>
  <AppDomains>
    <AppDomain>{frontend_url}</AppDomain>
    <AppDomain>{auth_proxy_url}</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name="Document"/>
    <Host Name="Workbook"/>
    <Host Name="Presentation"/>
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="{frontend_url}/taskpane.html?backend=streamassist"/>
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>

  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Document">
        <DesktopFormFactor>
          <GetStarted>
            <Title resid="GetStarted.Title"/>
            <Description resid="GetStarted.Description"/>
            <LearnMoreUrl resid="GetStarted.LearnMoreUrl"/>
          </GetStarted>
          <FunctionFile resid="Commands.Url"/>
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <OfficeTab id="TabHome">
              <Group id="CommandsGroupWordWIF">
                <Label resid="CommandsGroup.Label"/>
                <Icon>
                  <bt:Image size="16" resid="Icon.16x16"/>
                  <bt:Image size="32" resid="Icon.32x32"/>
                  <bt:Image size="80" resid="Icon.80x80"/>
                </Icon>
                <Control xsi:type="Button" id="TaskpaneButtonWordWIF">
                  <Label resid="TaskpaneButton.Label"/>
                  <Supertip>
                    <Title resid="TaskpaneButton.Label"/>
                    <Description resid="TaskpaneButton.Tooltip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16x16"/>
                    <bt:Image size="32" resid="Icon.32x32"/>
                    <bt:Image size="80" resid="Icon.80x80"/>
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>ButtonIdWordWIF</TaskpaneId>
                    <SourceLocation resid="Taskpane.Url"/>
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>

      <Host xsi:type="Workbook">
        <DesktopFormFactor>
          <GetStarted>
            <Title resid="GetStarted.Title"/>
            <Description resid="GetStarted.Description"/>
            <LearnMoreUrl resid="GetStarted.LearnMoreUrl"/>
          </GetStarted>
          <FunctionFile resid="Commands.Url"/>
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <OfficeTab id="TabHome">
              <Group id="CommandsGroupExcelWIF">
                <Label resid="CommandsGroup.Label"/>
                <Icon>
                  <bt:Image size="16" resid="Icon.16x16"/>
                  <bt:Image size="32" resid="Icon.32x32"/>
                  <bt:Image size="80" resid="Icon.80x80"/>
                </Icon>
                <Control xsi:type="Button" id="TaskpaneButtonExcelWIF">
                  <Label resid="TaskpaneButton.Label"/>
                  <Supertip>
                    <Title resid="TaskpaneButton.Label"/>
                    <Description resid="TaskpaneButton.Tooltip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16x16"/>
                    <bt:Image size="32" resid="Icon.32x32"/>
                    <bt:Image size="80" resid="Icon.80x80"/>
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>ButtonIdExcelWIF</TaskpaneId>
                    <SourceLocation resid="Taskpane.Url"/>
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>

      <Host xsi:type="Presentation">
        <DesktopFormFactor>
          <GetStarted>
            <Title resid="GetStarted.Title"/>
            <Description resid="GetStarted.Description"/>
            <LearnMoreUrl resid="GetStarted.LearnMoreUrl"/>
          </GetStarted>
          <FunctionFile resid="Commands.Url"/>
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <OfficeTab id="TabHome">
              <Group id="CommandsGroupPPTWIF">
                <Label resid="CommandsGroup.Label"/>
                <Icon>
                  <bt:Image size="16" resid="Icon.16x16"/>
                  <bt:Image size="32" resid="Icon.32x32"/>
                  <bt:Image size="80" resid="Icon.80x80"/>
                </Icon>
                <Control xsi:type="Button" id="TaskpaneButtonPPTWIF">
                  <Label resid="TaskpaneButton.Label"/>
                  <Supertip>
                    <Title resid="TaskpaneButton.Label"/>
                    <Description resid="TaskpaneButton.Tooltip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16x16"/>
                    <bt:Image size="32" resid="Icon.32x32"/>
                    <bt:Image size="80" resid="Icon.80x80"/>
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>ButtonIdPPTWIF</TaskpaneId>
                    <SourceLocation resid="Taskpane.Url"/>
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>

    <WebApplicationInfo>
      <Id>{client_id}</Id>
      <Resource>api://{frontend_domain}/{client_id}</Resource>
      <Scopes>
        <Scope>access_as_user</Scope>
      </Scopes>
    </WebApplicationInfo>

    <Resources>
      <bt:Images>
        <bt:Image id="Icon.16x16" DefaultValue="{frontend_url}/assets/icon-16.png"/>
        <bt:Image id="Icon.32x32" DefaultValue="{frontend_url}/assets/icon-32.png"/>
        <bt:Image id="Icon.80x80" DefaultValue="{frontend_url}/assets/icon-80.png"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="GetStarted.LearnMoreUrl" DefaultValue="https://go.microsoft.com/fwlink/?LinkId=276812"/>
        <bt:Url id="Commands.Url" DefaultValue="{frontend_url}/commands.html"/>
        <bt:Url id="Taskpane.Url" DefaultValue="{frontend_url}/taskpane.html?backend=streamassist"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="GetStarted.Title" DefaultValue="{display_name}"/>
        <bt:String id="CommandsGroup.Label" DefaultValue="{display_name}"/>
        <bt:String id="TaskpaneButton.Label" DefaultValue="{button_label}"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="GetStarted.Description" DefaultValue="{description}"/>
        <bt:String id="TaskpaneButton.Tooltip" DefaultValue="Open {display_name}."/>
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
"""

GSUITE_TEMPLATE = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">
  <Id>{manifest_id}</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>{provider_name}</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="{display_name}"/>
  <Description DefaultValue="{description}"/>
  <IconUrl DefaultValue="{frontend_url}/assets/icon-32.png"/>
  <HighResolutionIconUrl DefaultValue="{frontend_url}/assets/icon-64.png"/>
  <SupportUrl DefaultValue="https://cloud.google.com/vertex-ai"/>
  <AppDomains>
    <AppDomain>{frontend_url}</AppDomain>
    <AppDomain>{auth_proxy_url}</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name="Document"/>
    <Host Name="Workbook"/>
    <Host Name="Presentation"/>
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="{frontend_url}/taskpane.html"/>
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>

  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Document">
        <DesktopFormFactor>
          <GetStarted>
            <Title resid="GetStarted.Title"/>
            <Description resid="GetStarted.Description"/>
            <LearnMoreUrl resid="GetStarted.LearnMoreUrl"/>
          </GetStarted>
          <FunctionFile resid="Commands.Url"/>
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <OfficeTab id="TabHome">
              <Group id="CommandsGroupWordGSuite">
                <Label resid="CommandsGroup.Label"/>
                <Icon>
                  <bt:Image size="16" resid="Icon.16x16"/>
                  <bt:Image size="32" resid="Icon.32x32"/>
                  <bt:Image size="80" resid="Icon.80x80"/>
                </Icon>
                <Control xsi:type="Button" id="TaskpaneButtonWordGSuite">
                  <Label resid="TaskpaneButton.Label"/>
                  <Supertip>
                    <Title resid="TaskpaneButton.Label"/>
                    <Description resid="TaskpaneButton.Tooltip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16x16"/>
                    <bt:Image size="32" resid="Icon.32x32"/>
                    <bt:Image size="80" resid="Icon.80x80"/>
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>ButtonIdWordGSuite</TaskpaneId>
                    <SourceLocation resid="Taskpane.Url"/>
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>

      <Host xsi:type="Workbook">
        <DesktopFormFactor>
          <GetStarted>
            <Title resid="GetStarted.Title"/>
            <Description resid="GetStarted.Description"/>
            <LearnMoreUrl resid="GetStarted.LearnMoreUrl"/>
          </GetStarted>
          <FunctionFile resid="Commands.Url"/>
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <OfficeTab id="TabHome">
              <Group id="CommandsGroupExcelGSuite">
                <Label resid="CommandsGroup.Label"/>
                <Icon>
                  <bt:Image size="16" resid="Icon.16x16"/>
                  <bt:Image size="32" resid="Icon.32x32"/>
                  <bt:Image size="80" resid="Icon.80x80"/>
                </Icon>
                <Control xsi:type="Button" id="TaskpaneButtonExcelGSuite">
                  <Label resid="TaskpaneButton.Label"/>
                  <Supertip>
                    <Title resid="TaskpaneButton.Label"/>
                    <Description resid="TaskpaneButton.Tooltip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16x16"/>
                    <bt:Image size="32" resid="Icon.32x32"/>
                    <bt:Image size="80" resid="Icon.80x80"/>
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>ButtonIdExcelGSuite</TaskpaneId>
                    <SourceLocation resid="Taskpane.Url"/>
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>

      <Host xsi:type="Presentation">
        <DesktopFormFactor>
          <GetStarted>
            <Title resid="GetStarted.Title"/>
            <Description resid="GetStarted.Description"/>
            <LearnMoreUrl resid="GetStarted.LearnMoreUrl"/>
          </GetStarted>
          <FunctionFile resid="Commands.Url"/>
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <OfficeTab id="TabHome">
              <Group id="CommandsGroupPPTGSuite">
                <Label resid="CommandsGroup.Label"/>
                <Icon>
                  <bt:Image size="16" resid="Icon.16x16"/>
                  <bt:Image size="32" resid="Icon.32x32"/>
                  <bt:Image size="80" resid="Icon.80x80"/>
                </Icon>
                <Control xsi:type="Button" id="TaskpaneButtonPPTGSuite">
                  <Label resid="TaskpaneButton.Label"/>
                  <Supertip>
                    <Title resid="TaskpaneButton.Label"/>
                    <Description resid="TaskpaneButton.Tooltip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16x16"/>
                    <bt:Image size="32" resid="Icon.32x32"/>
                    <bt:Image size="80" resid="Icon.80x80"/>
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>ButtonIdPPTGSuite</TaskpaneId>
                    <SourceLocation resid="Taskpane.Url"/>
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>

    <WebApplicationInfo>
      <Id>{client_id}</Id>
      <Resource>api://{frontend_domain}/{client_id}</Resource>
      <Scopes>
        <Scope>access_as_user</Scope>
      </Scopes>
    </WebApplicationInfo>

    <Resources>
      <bt:Images>
        <bt:Image id="Icon.16x16" DefaultValue="{frontend_url}/assets/icon-16.png"/>
        <bt:Image id="Icon.32x32" DefaultValue="{frontend_url}/assets/icon-32.png"/>
        <bt:Image id="Icon.80x80" DefaultValue="{frontend_url}/assets/icon-80.png"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="GetStarted.LearnMoreUrl" DefaultValue="https://go.microsoft.com/fwlink/?LinkId=276812"/>
        <bt:Url id="Commands.Url" DefaultValue="{frontend_url}/commands.html"/>
        <bt:Url id="Taskpane.Url" DefaultValue="{frontend_url}/taskpane.html"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="GetStarted.Title" DefaultValue="{display_name}"/>
        <bt:String id="CommandsGroup.Label" DefaultValue="{display_name}"/>
        <bt:String id="TaskpaneButton.Label" DefaultValue="{button_label}"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="GetStarted.Description" DefaultValue="{description}"/>
        <bt:String id="TaskpaneButton.Tooltip" DefaultValue="Open {display_name}."/>
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
"""

def clean_url(url: str) -> str:
    url = url.strip().rstrip('/')
    if not url.startswith('http://') and not url.startswith('https://'):
        url = 'https://' + url
    return url

def extract_domain(url: str) -> str:
    cleaned = clean_url(url)
    return cleaned.replace('https://', '').replace('http://', '').split('/')[0]

def prompt_interactive():
    print("=" * 70)
    print("  🚀 Microsoft Office 365 Add-in Manifest Generator")
    print("  Author: Carlos Augusto, Principal Architect, Google")
    print("=" * 70)
    print()

    # 1. Track selection
    print("Choose your Deployment Track:")
    print("  1) Track 1: Workforce Identity Federation (WIF) - Seamless SSO")
    print("  2) Track 2: Cloud Identity / Google Workspace - 3-Legged OAuth")
    track_choice = input("Select [1 or 2] (default: 1): ").strip()
    track = "wif" if track_choice in ("", "1") else "gsuite"

    # 2. Frontend URL
    while True:
        frontend_input = input("\nEnter your Cloud Run gemini-frontend URL (e.g. https://gemini-frontend-xxx.run.app): ").strip()
        if frontend_input:
            frontend_url = clean_url(frontend_input)
            break
        print("❌ Error: gemini-frontend URL is required.")

    # 3. Auth Proxy URL (Optional, defaults to deriving or frontend domain)
    auth_proxy_input = input("Enter your Cloud Run auth-proxy URL (optional, press Enter to skip): ").strip()
    auth_proxy_url = clean_url(auth_proxy_input) if auth_proxy_input else frontend_url

    # 4. Microsoft Entra Client ID
    while True:
        client_id = input("\nEnter your Microsoft Entra Application (Client) ID: ").strip()
        if client_id:
            break
        print("❌ Error: Entra Client ID is required.")

    # 5. Display Name
    default_name = "Gemini Enterprise (WIF)" if track == "wif" else "Gemini Assistant (GSuite)"
    display_name_input = input(f"\nAdd-in Display Name (default: '{default_name}'): ").strip()
    display_name = display_name_input if display_name_input else default_name

    # 6. Button Label
    default_button = "Gemini (WIF)" if track == "wif" else "Gemini"
    button_input = input(f"Ribbon Button Label (default: '{default_button}'): ").strip()
    button_label = button_input if button_input else default_button

    # 7. Provider Name
    provider_input = input("Provider / Organization Name (default: 'Google Cloud'): ").strip()
    provider_name = provider_input if provider_input else "Google Cloud"

    # 8. Output filename
    default_out = f"manifest-{track}-custom.xml"
    out_input = input(f"\nOutput filename (default: '{default_out}'): ").strip()
    out_file = out_input if out_input else default_out

    return generate_manifest(
        track=track,
        frontend_url=frontend_url,
        auth_proxy_url=auth_proxy_url,
        client_id=client_id,
        display_name=display_name,
        button_label=button_label,
        provider_name=provider_name,
        out_file=out_file
    )

def generate_manifest(track, frontend_url, auth_proxy_url, client_id, display_name, button_label, provider_name, out_file):
    frontend_url = clean_url(frontend_url)
    auth_proxy_url = clean_url(auth_proxy_url)
    frontend_domain = extract_domain(frontend_url)
    manifest_id = str(uuid.uuid4())
    
    description = f"{display_name} for Microsoft Word, PowerPoint, and Excel."
    template = WIF_TEMPLATE if track == "wif" else GSUITE_TEMPLATE

    content = template.format(
        manifest_id=manifest_id,
        provider_name=provider_name,
        display_name=display_name,
        button_label=button_label,
        description=description,
        frontend_url=frontend_url,
        frontend_domain=frontend_domain,
        auth_proxy_url=auth_proxy_url,
        client_id=client_id
    )

    with open(out_file, "w", encoding="utf-8") as f:
        f.write(content.strip() + "\n")

    print()
    print("=" * 70)
    print(f"  ✅ Manifest successfully created: {out_file}")
    print("=" * 70)
    print(f"  • Unique Manifest GUID: {manifest_id}")
    print(f"  • Entra App Client ID:  {client_id}")
    print(f"  • Application ID URI:   api://{frontend_domain}/{client_id}")
    print(f"  • Taskpane Source:      {frontend_url}/taskpane.html")
    print("=" * 70)
    print("Next step: Sideload or deploy this manifest in the Microsoft 365 Admin Center.")
    return out_file

def main():
    parser = argparse.ArgumentParser(description="Generate Microsoft Office Add-in XML Manifest for Gemini Enterprise.")
    parser.add_argument("--track", choices=["wif", "gsuite"], help="Deployment track: 'wif' or 'gsuite'")
    parser.add_argument("--frontend-url", help="Public HTTPS URL of gemini-frontend Cloud Run service")
    parser.add_argument("--auth-proxy-url", help="Public HTTPS URL of auth-proxy Cloud Run service")
    parser.add_argument("--client-id", help="Microsoft Entra Application (Client) ID")
    parser.add_argument("--display-name", help="Add-in Display Name on Ribbon")
    parser.add_argument("--button-label", help="Ribbon Button Label")
    parser.add_argument("--provider-name", default="Google Cloud", help="Provider/Organization Name")
    parser.add_argument("--output", help="Output file path (e.g. manifest-custom.xml)")

    args = parser.parse_args()

    if len(sys.argv) == 1:
        # Interactive mode
        prompt_interactive()
    else:
        if not args.frontend_url or not args.client_id:
            print("❌ Error: --frontend-url and --client-id are required in non-interactive mode.", file=sys.stderr)
            sys.exit(1)
        
        track = args.track or "wif"
        display_name = args.display_name or ("Gemini Enterprise (WIF)" if track == "wif" else "Gemini Assistant (GSuite)")
        button_label = args.button_label or ("Gemini (WIF)" if track == "wif" else "Gemini")
        auth_proxy_url = args.auth_proxy_url or args.frontend_url
        out_file = args.output or f"manifest-{track}-custom.xml"

        generate_manifest(
            track=track,
            frontend_url=args.frontend_url,
            auth_proxy_url=auth_proxy_url,
            client_id=args.client_id,
            display_name=display_name,
            button_label=button_label,
            provider_name=args.provider_name,
            out_file=out_file
        )

if __name__ == "__main__":
    main()
