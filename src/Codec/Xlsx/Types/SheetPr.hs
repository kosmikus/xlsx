{-# LANGUAGE CPP #-}
{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE TemplateHaskell   #-}
{-# LANGUAGE RecordWildCards   #-}
{-# LANGUAGE DeriveGeneric #-}
module Codec.Xlsx.Types.SheetPr (
    -- * Main types
    SheetPr(..)
    -- * Lenses
    -- ** SheetPr
  , sheetPrTabColor
  , sheetPrOutlinePr
  , sheetPrPageSetUpPr
  , sheetPrSyncHorizontal
  , sheetPrSyncVertical
  , sheetPrSyncRef
  , sheetPrTransitionEvaluation
  , sheetPrTransitionEntry
  , sheetPrPublished
  , sheetPrCodeName
  , sheetPrFilterMode
  , sheetPrEnableFormatConditionsCalculation
  , outlinePrApplyStyles
  , outlinePrSummaryBelow
  , outlinePrSummaryRight
  , outlinePrShowOutlineSymbols
  , pageSetUpPrAutoPageBreaks
  , pageSetUpPrFitToPage
  ) where

#ifdef USE_MICROLENS
import Lens.Micro.TH (makeLenses)
#else
import Control.Lens (makeLenses)
#endif
import Control.DeepSeq (NFData)
import Data.Default
import qualified Data.Map as Map
import Data.Maybe (catMaybes)
import Data.Text (Text)
import GHC.Generics (Generic)
import Text.XML

import Codec.Xlsx.Types.StyleSheet
import Codec.Xlsx.Writer.Internal
import Codec.Xlsx.Parser.Internal

{-------------------------------------------------------------------------------
  Main types
-------------------------------------------------------------------------------}

data SheetPr = SheetPr {
    _sheetPrTabColor :: Maybe Color
  , _sheetPrOutlinePr :: Maybe OutlinePr
  , _sheetPrPageSetUpPr :: Maybe PageSetUpPr
  , _sheetPrSyncHorizontal :: Maybe Bool
  , _sheetPrSyncVertical :: Maybe Bool
  , _sheetPrSyncRef :: Maybe Bool
  , _sheetPrTransitionEvaluation :: Maybe Bool
  , _sheetPrTransitionEntry :: Maybe Bool
  , _sheetPrPublished :: Maybe Bool
  , _sheetPrCodeName :: Maybe Text
  , _sheetPrFilterMode :: Maybe Bool
  , _sheetPrEnableFormatConditionsCalculation :: Maybe Bool
  }
  deriving (Eq, Ord, Show, Generic)
instance NFData SheetPr

data OutlinePr = OutlinePr {
    _outlinePrApplyStyles :: Maybe Bool
  , _outlinePrSummaryBelow :: Maybe Bool
  , _outlinePrSummaryRight :: Maybe Bool
  , _outlinePrShowOutlineSymbols :: Maybe Bool
  }
  deriving (Eq, Ord, Show, Generic)
instance NFData OutlinePr

data PageSetUpPr = PageSetUpPr {
    _pageSetUpPrAutoPageBreaks :: Maybe Bool
  , _pageSetUpPrFitToPage :: Maybe Bool
  }
  deriving (Eq, Ord, Show, Generic)
instance NFData PageSetUpPr

{-------------------------------------------------------------------------------
  Default instances
-------------------------------------------------------------------------------}

instance Default SheetPr where
  def = SheetPr {
      _sheetPrTabColor = Nothing
    , _sheetPrOutlinePr = Nothing
    , _sheetPrPageSetUpPr = Nothing
    , _sheetPrSyncHorizontal = Nothing
    , _sheetPrSyncVertical = Nothing
    , _sheetPrSyncRef = Nothing
    , _sheetPrTransitionEvaluation = Nothing
    , _sheetPrTransitionEntry = Nothing
    , _sheetPrPublished = Nothing
    , _sheetPrCodeName = Nothing
    , _sheetPrFilterMode = Nothing
    , _sheetPrEnableFormatConditionsCalculation = Nothing
    }

instance Default OutlinePr where
  def = OutlinePr {
      _outlinePrApplyStyles = Nothing
    , _outlinePrSummaryBelow = Nothing
    , _outlinePrSummaryRight = Nothing
    , _outlinePrShowOutlineSymbols = Nothing
    }

instance Default PageSetUpPr where
  def = PageSetUpPr {
      _pageSetUpPrAutoPageBreaks = Nothing
    , _pageSetUpPrFitToPage = Nothing
    }

{-------------------------------------------------------------------------------
  Lenses
-------------------------------------------------------------------------------}

makeLenses ''SheetPr
makeLenses ''OutlinePr
makeLenses ''PageSetUpPr

{-------------------------------------------------------------------------------
  Rendering
-------------------------------------------------------------------------------}

-- | See @CT_PageSetup@, p. 3922
instance ToElement SheetPr where
  toElement nm SheetPr{..} = Element {
      elementName       = nm
    , elementNodes      = map NodeElement . catMaybes $ [
          toElement "tabColor" <$> _sheetPrTabColor
        , toElement "outlinePr" <$> _sheetPrOutlinePr
        , toElement "pageSetUpPr" <$> _sheetPrPageSetUpPr
        ]
    , elementAttributes = Map.fromList . catMaybes $ [
          "syncHorizontal" .=? _sheetPrSyncHorizontal
        , "syncVertical" .=? _sheetPrSyncVertical
        , "syncRef" .=? _sheetPrSyncRef
        , "transitionEvaluation" .=? _sheetPrTransitionEvaluation
        , "transitionEntry" .=? _sheetPrTransitionEntry
        , "published" .=? _sheetPrPublished
        , "codeName" .=? _sheetPrCodeName
        , "filterMode" .=? _sheetPrFilterMode
        , "enableFormatConditionsCalculation" .=? _sheetPrEnableFormatConditionsCalculation
        ]
    }

instance ToElement OutlinePr where
  toElement nm OutlinePr{..} = Element {
      elementName       = nm
    , elementNodes      = []
    , elementAttributes = Map.fromList . catMaybes $ [
          "applyStyles" .=? _outlinePrApplyStyles
        , "summaryBelow" .=? _outlinePrSummaryBelow
        , "summaryRight" .=? _outlinePrSummaryRight
        , "showOutlineSymbols" .=? _outlinePrShowOutlineSymbols
        ]
    }

instance ToElement PageSetUpPr where
  toElement nm PageSetUpPr{..} = Element {
      elementName       = nm
    , elementNodes      = []
    , elementAttributes = Map.fromList . catMaybes $ [
          "autoPageBreaks" .=? _pageSetUpPrAutoPageBreaks
        , "fitToPage"      .=? _pageSetUpPrFitToPage
        ]
    }


{-------------------------------------------------------------------------------
  Parsing
-------------------------------------------------------------------------------}

instance FromCursor SheetPr where
  fromCursor cur = do
    _sheetPrTabColor                          <- maybeFromElement (n_ "tabColor") cur
    _sheetPrOutlinePr                         <- maybeFromElement (n_ "outlinePr") cur
    _sheetPrPageSetUpPr                       <- maybeFromElement (n_ "pageSetUpPr") cur
    _sheetPrSyncHorizontal                    <- maybeAttribute "syncHorizontal" cur
    _sheetPrSyncVertical                      <- maybeAttribute "syncVertical" cur
    _sheetPrSyncRef                           <- maybeAttribute "syncRef" cur
    _sheetPrTransitionEvaluation              <- maybeAttribute "transitionEvaluation" cur
    _sheetPrTransitionEntry                   <- maybeAttribute "transitionEntry" cur
    _sheetPrPublished                         <- maybeAttribute "published" cur
    _sheetPrCodeName                          <- maybeAttribute "codeName" cur
    _sheetPrFilterMode                        <- maybeAttribute "filterMode" cur
    _sheetPrEnableFormatConditionsCalculation <- maybeAttribute "enableFormatConditionsCalculation" cur
    return SheetPr{..}

instance FromCursor OutlinePr where
  fromCursor cur = do
    _outlinePrApplyStyles        <- maybeAttribute "applyStyles" cur
    _outlinePrSummaryBelow       <- maybeAttribute "summaryBelow" cur
    _outlinePrSummaryRight       <- maybeAttribute "summaryRight" cur
    _outlinePrShowOutlineSymbols <- maybeAttribute "showOutlineSymbols" cur
    return OutlinePr{..}

instance FromCursor PageSetUpPr where
  fromCursor cur = do
    _pageSetUpPrAutoPageBreaks <- maybeAttribute "autoPageBreaks" cur
    _pageSetUpPrFitToPage      <- maybeAttribute "fitToPage" cur
    return PageSetUpPr{..}

instance FromXenoNode SheetPr where
  fromXenoNode root = do
    c <- collectChildren root $ do
      _sheetPrTabColor    <- maybeFromChild "tabColor"
      _sheetPrOutlinePr   <- maybeFromChild "outlinePr"
      _sheetPrPageSetUpPr <- maybeFromChild "pageSetUpPr"
      return (SheetPr _sheetPrTabColor _sheetPrOutlinePr _sheetPrPageSetUpPr)
    parseAttributes root $ do
      _sheetPrSyncHorizontal                    <- maybeAttr "syncHorizontal"
      _sheetPrSyncVertical                      <- maybeAttr "syncVertical"
      _sheetPrSyncRef                           <- maybeAttr "syncRef"
      _sheetPrTransitionEvaluation              <- maybeAttr "transitionEvaluation"
      _sheetPrTransitionEntry                   <- maybeAttr "transitionEntry"
      _sheetPrPublished                         <- maybeAttr "published"
      _sheetPrCodeName                          <- maybeAttr "codeName"
      _sheetPrFilterMode                        <- maybeAttr "filterMode"
      _sheetPrEnableFormatConditionsCalculation <- maybeAttr "enableFormatConditionsCalculation"
      return $
        c
          _sheetPrSyncHorizontal
          _sheetPrSyncVertical
          _sheetPrSyncRef
          _sheetPrTransitionEvaluation
          _sheetPrTransitionEntry
          _sheetPrPublished
          _sheetPrCodeName
          _sheetPrFilterMode
          _sheetPrEnableFormatConditionsCalculation

instance FromXenoNode OutlinePr where
  fromXenoNode root =
    parseAttributes root $ do
      _outlinePrApplyStyles        <- maybeAttr "applyStyles"
      _outlinePrSummaryBelow       <- maybeAttr "summaryBelow"
      _outlinePrSummaryRight       <- maybeAttr "summaryRight"
      _outlinePrShowOutlineSymbols <- maybeAttr "showOutlineSymbols"
      return OutlinePr{..}

instance FromXenoNode PageSetUpPr where
  fromXenoNode root =
    parseAttributes root $ do
      _pageSetUpPrAutoPageBreaks <- maybeAttr "autoPageBreaks"
      _pageSetUpPrFitToPage      <- maybeAttr "fitToPage"
      return PageSetUpPr {..}
