"""
Unit & Integration Tests for Auth-Proxy FastAPI Microservice
"""
import os
import unittest
from unittest.mock import patch
from fastapi.testclient import TestClient

# Set test environment defaults before importing app
os.environ["REQUIRE_ENTRA_AUTH"] = "false"
os.environ["MICROSOFT_ENTRA_APP_ID"] = "test-client-guid-12345"

from main import app, extract_user_from_payload, AuthenticatedUser

client = TestClient(app)


class TestAuthProxy(unittest.TestCase):

    def test_health_check(self):
        """Verify health check endpoints return 200 OK and expected structure."""
        response = client.get("/health")
        self.assertEqual(response.status_code, 200)
        data = response.json()
        self.assertEqual(data["status"], "healthy")
        self.assertEqual(data["service"], "auth-proxy")
        self.assertTrue(data["entra_app_id_configured"])
        self.assertIn("verbose_logging", data)

    def test_root_endpoint(self):
        """Verify root / endpoint maps to health check."""
        response = client.get("/")
        self.assertEqual(response.status_code, 200)
        data = response.json()
        self.assertEqual(data["status"], "healthy")
        self.assertIn("verbose_logging", data)

    def test_extract_user_from_payload(self):
        """Verify claim extraction helper properly captures user identity."""
        mock_payload = {
            "preferred_username": "john.doe@contoso.com",
            "name": "John Doe",
            "tid": "tenant-abc-123",
            "oid": "user-oid-456",
            "sub": "subject-789",
            "scp": "access_as_user profile email",
            "roles": ["Admin", "PowerUser"]
        }
        user = extract_user_from_payload(mock_payload)
        self.assertEqual(user.user_id, "john.doe@contoso.com")
        self.assertEqual(user.email, "john.doe@contoso.com")
        self.assertEqual(user.name, "John Doe")
        self.assertEqual(user.tenant_id, "tenant-abc-123")
        self.assertEqual(user.oid, "user-oid-456")
        self.assertIn("access_as_user", user.scopes)
        self.assertIn("Admin", user.roles)

    def test_extract_user_from_payload_with_raw_token(self):
        """Verify claim extraction helper properly captures user identity and raw token."""
        mock_payload = {
            "preferred_username": "alexw@5m4qby.onmicrosoft.com",
            "name": "Alex Wilber",
            "tid": "tenant-abc-123",
            "oid": "user-oid-456",
            "sub": "subject-789"
        }
        user = extract_user_from_payload(mock_payload, raw_token="mock_raw_jwt_token_123")
        self.assertEqual(user.user_id, "alexw@5m4qby.onmicrosoft.com")
        self.assertEqual(user.raw_token, "mock_raw_jwt_token_123")

    @patch("main.get_discovery_engine_acl_config")
    def test_resolve_end_user_google_token_gsuite(self, mock_acl):
        """Verify resolve_end_user_google_token correctly determines cloud_identity mode for GSUITE projects."""
        mock_acl.return_value = {"idpConfig": {"idpType": "GSUITE"}}
        from main import resolve_end_user_google_token
        user = AuthenticatedUser(user_id="alexw@5m4qby.onmicrosoft.com", raw_token="mock_jwt")
        token, mode, diag = resolve_end_user_google_token(user, project_id="agentspace-452714", location="global")
        self.assertEqual(mode, "cloud_identity")
        self.assertIn("discovered_idp_type", diag)
        self.assertEqual(diag["discovered_idp_type"], "GSUITE")

    @patch("main.get_discovery_engine_acl_config")
    def test_resolve_end_user_google_token_wif(self, mock_acl):
        """Verify resolve_end_user_google_token correctly determines wif mode for THIRD_PARTY projects."""
        mock_acl.return_value = {
            "idpConfig": {
                "idpType": "THIRD_PARTY",
                "externalIdpConfig": {
                    "workforcePoolName": "locations/global/workforcePools/ca-entra-id-oidc-pool"
                }
            }
        }
        from main import resolve_end_user_google_token
        user = AuthenticatedUser(user_id="alexw@5m4qby.onmicrosoft.com", raw_token="mock_jwt")
        token, mode, diag = resolve_end_user_google_token(user, project_id="agentspace-wif", location="global")
        self.assertEqual(mode, "wif")
        self.assertIn("workforce_pool", diag)
        self.assertEqual(diag["workforce_pool"], "locations/global/workforcePools/ca-entra-id-oidc-pool")


if __name__ == "__main__":
    unittest.main()

