# Security Policy

## Supported versions

| Version | Supported |
|---|---|
| The current stable release (npm `latest`) | ✅ Fixes and security fixes |
| The previous minor release | ✅ Critical and security fixes for 30 days after the next minor ships (npm `support-X.Y`) |
| Betas (npm `beta`) | Fixed in the next beta |
| Anything older | ❌ |

1.4.x keeps critical and security fixes for 30 days after 1.5.0 ships, because 1.5 requires Power Platform ToolBox 1.2.7. See [RELEASING.md](RELEASING.md) → *Supported versions*.

## Reporting a Vulnerability

Please do **not** open a public issue.
Use GitHub Security Advisories: <https://github.com/mohsinonxrm/pptb-mxrm-fetchxml-studio/security/advisories/new>

## Scope

- Authentication flows
- Token handling
- Dataverse request logging
- Dependency vulnerabilities
