using MemoQ.MTInterfaces;
using MultiSupplierMTPlugin.Helpers;
using System;
using System.Drawing;
using System.Reflection;

namespace MultiSupplierMTPlugin
{
    class MultiSupplierMTEngine : EngineBase
    {
        private readonly MultiSupplierMTOptions _mtOptions;

        private readonly LimitHelper _limitHelper;

        private readonly RetryHelper _retryHelper;

        private readonly MultiSupplierMTService _providerService;

        private readonly RequestType _requestType;

        private readonly string _srcLangCode;

        private readonly string _trgLangCode;

        public MultiSupplierMTEngine(MultiSupplierMTOptions mtOptions, LimitHelper rateLimitHelper, RetryHelper retryHelper,
            MultiSupplierMTService providerService, RequestType _requestType, string srcLangCode, string trgLangCode)
        {
            this._mtOptions = mtOptions;

            this._limitHelper = rateLimitHelper;
            this._retryHelper = retryHelper;

            this._providerService = providerService;
            this._requestType = _requestType;

            this._srcLangCode = srcLangCode;
            this._trgLangCode = trgLangCode;
        }

        #region IEngine Members

        /// <summary>
        /// Tells if the engine supports the adjustment of fuzzy TM hits through machine translation (MatchPatch).
        /// This means that if there is a TM match for the source segment, but it is not perfect,
        /// memoQ will try to improve the suggestion by sending the difference to an MT provider for translation.
        /// If your MT service can only translate complete segments reliably, but not partial ones (e.g., two separate words),
        /// disable this feature. But if the service is good at translating segment parts, enable it. If the feature
        /// is disabled, your plugin will not appear in the MatchPatch list on the Edit machine translation
        /// settings dialog's Settings tab
        /// </summary>
        public override bool SupportsFuzzyCorrection
        {
            get { return true; }
        }

        public override void SetProperty(string name, string value)
        {
            throw new NotImplementedException();
        }

        public override Image SmallIcon
        {
            get
            {
                return Image.FromStream(Assembly.GetExecutingAssembly().GetManifestResourceStream("MultiSupplierMTPlugin.Icon.png"));
            }
        }

        public override ISession CreateLookupSession()
        {
            try
            {
                LoggingHelper.LogForсe($"MultiSupplierMTEngine|CreateLookupSession| _srcLangCode - {_srcLangCode}, _trgLangCode - {_trgLangCode}");
                return new MultiSupplierMTSession(_mtOptions, _limitHelper, _retryHelper, _providerService, _requestType, _srcLangCode, _trgLangCode);
            }
            catch (Exception e)
            {
                LoggingHelper.LogForсe($"MultiSupplierMTSession|CreateLookupSession| e.Message - {e.Message}, e.StackTrace - {e.StackTrace}");
                throw;
            }
        }

        public override ISessionForStoringTranslations CreateStoreTranslationSession()
        {
            try
            {
                LoggingHelper.LogForсe($"MultiSupplierMTEngine|CreateStoreTranslationSession| _srcLangCode - {_srcLangCode}, _trgLangCode - {_trgLangCode}");
                return new MultiSupplierMTSession(_mtOptions, _limitHelper, _retryHelper, _providerService, _requestType, _srcLangCode, _trgLangCode);
            }
            catch (Exception e)
            {
                LoggingHelper.LogForсe($"MultiSupplierMTSession|CreateStoreTranslationSession| e.Message - {e.Message}, e.StackTrace - {e.StackTrace}");
                throw;
            }
        }

        #endregion

        #region IDisposable Members

        public override void Dispose()
        {
        }

        #endregion
    }
}
