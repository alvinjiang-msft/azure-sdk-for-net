// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text.Json;

namespace Azure.Communication
{
    internal class CommunicationIdentifierSerializer
    {
        private const string CommunicationUserKindValue = "communicationUser";
        private const string PhoneNumberKindValue = "phoneNumber";
        private const string MicrosoftTeamsUserKindValue = "microsoftTeamsUser";
        private const string MicrosoftTeamsAppKindValue = "microsoftTeamsApp";
        private const string TeamsExtensionUserKindValue = "teamsExtensionUser";
        private const string TeamsExtensionUserPropertyName = "TeamsExtensionUser";
        private const string TeamsExtensionUserModelName = "TeamsExtensionUserIdentifierModel";

        public static CommunicationIdentifier Deserialize(CommunicationIdentifierModel identifier)
        {
            string rawId = AssertNotNull(identifier.RawId, nameof(identifier.RawId), nameof(CommunicationIdentifierModel));

            AssertMaximumOneNestedModel(identifier);

            var kind = identifier.Kind?.ToString() ?? GetKindValue(identifier);

            if (string.Equals(kind, CommunicationUserKindValue, StringComparison.OrdinalIgnoreCase)
                && identifier.CommunicationUser is not null)
            {
                return new CommunicationUserIdentifier(AssertNotNull(identifier.CommunicationUser.Id, nameof(identifier.CommunicationUser.Id), nameof(CommunicationUserIdentifierModel)));
            }

            if (string.Equals(kind, PhoneNumberKindValue, StringComparison.OrdinalIgnoreCase)
                && identifier.PhoneNumber is not null)
            {
                return new PhoneNumberIdentifier(
                    AssertNotNull(identifier.PhoneNumber.Value, nameof(identifier.PhoneNumber.Value), nameof(PhoneNumberIdentifierModel)),
                    AssertNotNull(identifier.RawId, nameof(identifier.RawId), nameof(PhoneNumberIdentifierModel)));
            }

            if (string.Equals(kind, MicrosoftTeamsUserKindValue, StringComparison.OrdinalIgnoreCase)
                && identifier.MicrosoftTeamsUser is not null)
            {
                var user = identifier.MicrosoftTeamsUser;
                return new MicrosoftTeamsUserIdentifier(
                      AssertNotNull(user.UserId, nameof(user.UserId), nameof(MicrosoftTeamsUserIdentifierModel)),
                      AssertNotNull(user.IsAnonymous, nameof(user.IsAnonymous), nameof(MicrosoftTeamsUserIdentifierModel)),
                      Deserialize(AssertNotNull(user.Cloud, nameof(user.Cloud), nameof(MicrosoftTeamsUserIdentifierModel))),
                      rawId);
            }

            if (string.Equals(kind, MicrosoftTeamsAppKindValue, StringComparison.OrdinalIgnoreCase)
                 && identifier.MicrosoftTeamsApp is not null)
            {
                var app = identifier.MicrosoftTeamsApp;
                return new MicrosoftTeamsAppIdentifier(
                      AssertNotNull(app.AppId, nameof(app.AppId), nameof(MicrosoftTeamsAppIdentifierModel)),
                      Deserialize(AssertNotNull(app.Cloud, nameof(app.Cloud), nameof(MicrosoftTeamsAppIdentifierModel))));
            }

            object teamsExtensionUser = GetPropertyValue(identifier, TeamsExtensionUserPropertyName);
            if (string.Equals(kind, TeamsExtensionUserKindValue, StringComparison.OrdinalIgnoreCase)
                && teamsExtensionUser is not null)
            {
                return new TeamsExtensionUserIdentifier(
                    AssertNotNull((string)GetPropertyValue(teamsExtensionUser, "UserId"), "UserId", TeamsExtensionUserModelName),
                    AssertNotNull((string)GetPropertyValue(teamsExtensionUser, "TenantId"), "TenantId", TeamsExtensionUserModelName),
                    AssertNotNull((string)GetPropertyValue(teamsExtensionUser, "ResourceId"), "ResourceId", TeamsExtensionUserModelName),
                    Deserialize(AssertNotNull((CommunicationCloudEnvironmentModel?)GetPropertyValue(teamsExtensionUser, "Cloud"), "Cloud", TeamsExtensionUserModelName)),
                    rawId);
            }

            return new UnknownIdentifier(rawId);

            static void AssertMaximumOneNestedModel(CommunicationIdentifierModel identifier)
            {
                List<string> presentProperties = new();
                if (identifier.CommunicationUser is not null)
                    presentProperties.Add(nameof(identifier.CommunicationUser));
                if (identifier.PhoneNumber is not null)
                    presentProperties.Add(nameof(identifier.PhoneNumber));
                if (identifier.MicrosoftTeamsUser is not null)
                    presentProperties.Add(nameof(identifier.MicrosoftTeamsUser));
                if (identifier.MicrosoftTeamsApp is not null)
                    presentProperties.Add(nameof(identifier.MicrosoftTeamsApp));
                if (GetPropertyValue(identifier, TeamsExtensionUserPropertyName) is not null)
                    presentProperties.Add(TeamsExtensionUserPropertyName);

                if (presentProperties.Count > 1)
                    throw new JsonException($"Only one of the properties in {{{string.Join(", ", presentProperties)}}} should be present.");
            }
        }

        internal static CommunicationIdentifierModelKind GetKind(CommunicationIdentifierModel identifier)
            => new CommunicationIdentifierModelKind(GetKindValue(identifier));

        internal static string GetKindValue(CommunicationIdentifierModel identifier)
        {
            if (identifier.CommunicationUser is not null)
            {
                return CommunicationUserKindValue;
            }

            if (identifier.PhoneNumber is not null)
            {
                return PhoneNumberKindValue;
            }

            if (identifier.MicrosoftTeamsUser is not null)
            {
                return MicrosoftTeamsUserKindValue;
            }

            if (identifier.MicrosoftTeamsApp is not null)
            {
                return MicrosoftTeamsAppKindValue;
            }

            if (GetPropertyValue(identifier, TeamsExtensionUserPropertyName) is not null)
            {
                return TeamsExtensionUserKindValue;
            }

            return CommunicationIdentifierModelKind.Unknown.ToString();
        }

        internal static CommunicationCloudEnvironment Deserialize(CommunicationCloudEnvironmentModel cloud)
        {
            if (cloud == CommunicationCloudEnvironmentModel.Public)
                return CommunicationCloudEnvironment.Public;
            if (cloud == CommunicationCloudEnvironmentModel.Gcch)
                return CommunicationCloudEnvironment.Gcch;
            if (cloud == CommunicationCloudEnvironmentModel.Dod)
                return CommunicationCloudEnvironment.Dod;

            return new CommunicationCloudEnvironment(cloud.ToString());
        }

        public static CommunicationIdentifierModel Serialize(CommunicationIdentifier identifier)
            => identifier switch
            {
                CommunicationUserIdentifier u => new CommunicationIdentifierModel
                {
                    RawId = u.Id,
                    CommunicationUser = new CommunicationUserIdentifierModel(u.Id),
                },
                PhoneNumberIdentifier p => SerializePhoneNumber(p),
                MicrosoftTeamsUserIdentifier u => new CommunicationIdentifierModel
                {
                    RawId = u.RawId,
                    MicrosoftTeamsUser = new MicrosoftTeamsUserIdentifierModel(u.UserId)
                    {
                        IsAnonymous = u.IsAnonymous,
                        Cloud = Serialize(u.Cloud),
                    }
                },
                MicrosoftTeamsAppIdentifier app => new CommunicationIdentifierModel
                {
                    RawId = app.RawId,
                    MicrosoftTeamsApp = new MicrosoftTeamsAppIdentifierModel(app.AppId)
                    {
                        Cloud = Serialize(app.Cloud),
                    }
                },
                TeamsExtensionUserIdentifier user => SerializeTeamsExtensionUser(user),
                UnknownIdentifier u => new CommunicationIdentifierModel
                {
                    RawId = u.Id
                },
                _ => throw new NotSupportedException(),
            };

        private static CommunicationIdentifierModel SerializePhoneNumber(PhoneNumberIdentifier identifier)
        {
            var phoneNumber = new PhoneNumberIdentifierModel(identifier.PhoneNumber);
            SetPropertyValue(phoneNumber, "IsAnonymous", identifier.IsAnonymous);
            SetPropertyValue(phoneNumber, "AssertedId", identifier.AssertedId);

            return new CommunicationIdentifierModel
            {
                RawId = identifier.RawId,
                PhoneNumber = phoneNumber,
            };
        }

        private static CommunicationIdentifierModel SerializeTeamsExtensionUser(TeamsExtensionUserIdentifier identifier)
        {
            var model = new CommunicationIdentifierModel
            {
                RawId = identifier.RawId,
            };

            object teamsExtensionUser = CreatePropertyInstance(model, TeamsExtensionUserPropertyName, identifier.UserId, identifier.TenantId, identifier.ResourceId);
            if (teamsExtensionUser is null)
            {
                throw new NotSupportedException($"{nameof(TeamsExtensionUserIdentifier)} is not supported by this package version.");
            }

            SetPropertyValue(teamsExtensionUser, "Cloud", Serialize(identifier.Cloud));
            SetPropertyValue(model, TeamsExtensionUserPropertyName, teamsExtensionUser);
            return model;
        }

        internal static CommunicationCloudEnvironmentModel Serialize(CommunicationCloudEnvironment cloud)
        {
            if (cloud == CommunicationCloudEnvironment.Public)
                return CommunicationCloudEnvironmentModel.Public;
            if (cloud == CommunicationCloudEnvironment.Gcch)
                return CommunicationCloudEnvironmentModel.Gcch;
            if (cloud == CommunicationCloudEnvironment.Dod)
                return CommunicationCloudEnvironmentModel.Dod;

            return new CommunicationCloudEnvironmentModel(cloud.ToString());
        }

        internal static T AssertNotNull<T>(T value, string name, string type) where T : class
            => value ?? throw new JsonException($"Property '{name}' is required for identifier of type `{type}`.");

        internal static T AssertNotNull<T>(T? value, string name, string type) where T : struct
        {
            if (value is null)
                throw new JsonException($"Property '{name}' is required for identifier of type `{type}`.");

            return value.Value;
        }

        private static object GetPropertyValue(object instance, string propertyName)
            => instance.GetType().GetProperty(propertyName, BindingFlags.Instance | BindingFlags.Public)?.GetValue(instance);

        private static void SetPropertyValue(object instance, string propertyName, object value)
        {
            PropertyInfo property = instance.GetType().GetProperty(propertyName, BindingFlags.Instance | BindingFlags.Public);
            property?.SetValue(instance, value);
        }

        private static object CreatePropertyInstance(object instance, string propertyName, params object[] constructorArguments)
        {
            PropertyInfo property = instance.GetType().GetProperty(propertyName, BindingFlags.Instance | BindingFlags.Public);
            return property is null ? null : Activator.CreateInstance(property.PropertyType, constructorArguments);
        }
    }
}
